//! End-to-end language tests: source in, results out.
#![allow(clippy::unwrap_used, clippy::panic)]

use macroost_lang::ast::ModuleType;
use macroost_lang::{Interpreter, NoHost, Value};

fn run_with(modules: &[(&str, ModuleType, &str)], entry: &str) -> (Result<Value, macroost_lang::RuntimeError>, NoHost) {
    let mut host = NoHost::default();
    let r = {
        let mut it = Interpreter::new(&mut host);
        for (n, k, s) in modules {
            it.add_module(n, *k, s).unwrap();
        }
        it.run(entry, vec![])
    };
    (r, host)
}

fn run(src: &str, entry: &str) -> Value {
    let (r, _) = run_with(&[("Module1", ModuleType::Standard, src)], entry);
    r.unwrap()
}

fn text(v: Value) -> String {
    v.to_text().unwrap()
}

#[test]
fn functions_loops_and_byref() {
    let src = r#"
Function Twice(ByVal n As Long) As Long
    Twice = n * 2
End Function
Sub Bump(x As Long)
    x = x + 1
End Sub
Sub NoBump(ByVal x As Long)
    x = x + 1
End Sub
Function T() As String
    Dim i As Long, total As Long, a As Long, b As Long
    For i = 1 To 10
        total = total + i
    Next i
    a = 5: b = 5
    Bump a
    NoBump b
    Call Bump(a)
    T = total & "," & Twice(21) & "," & a & "," & b
End Function"#;
    assert_eq!(text(run(src, "T")), "55,42,7,5");
}

#[test]
fn recursion_and_exit() {
    let src = r#"
Function Fact(n As Long) As Long
    If n <= 1 Then
        Fact = 1
        Exit Function
    End If
    Fact = n * Fact(n - 1)
End Function
Function T()
    T = Fact(10)
End Function"#;
    assert_eq!(text(run(src, "T")), "3628800");
}

#[test]
fn arrays_redim_preserve_and_for_each() {
    let src = r#"
Function T() As String
    Dim a() As Long, i As Long, s As String, v
    ReDim a(1 To 3)
    For i = 1 To 3: a(i) = i * 10: Next
    ReDim Preserve a(1 To 4)
    a(4) = 40
    For Each v In a
        s = s & v & " "
    Next
    Dim m(2, 3) As Integer
    m(2, 3) = 7
    T = s & m(2, 3)
End Function"#;
    assert_eq!(text(run(src, "T")), "10 20 30 40 7");
}

#[test]
fn do_while_select_and_with() {
    let src = r#"
Function Grade(n As Integer) As String
    Select Case n
        Case Is > 89: Grade = "A"
        Case 80 To 89: Grade = "B"
        Case 70, 71, 72: Grade = "C-"
        Case Else: Grade = "F"
    End Select
End Function
Function T() As String
    Dim k As Integer, s As String
    Do While k < 3
        k = k + 1
    Loop
    Do
        k = k + 1
    Loop Until k = 6
    Dim c As New Collection
    With c
        .Add "x"
        .Add "y", "key"
    End With
    s = Grade(95) & Grade(85) & Grade(71) & Grade(10)
    T = s & k & c.Count & c("key")
End Function"#;
    assert_eq!(text(run(src, "T")), "ABC-F62y");
}

#[test]
fn error_handling_resume_next_and_goto() {
    let src = r#"
Function T() As String
    Dim c As New Collection, s As String
    On Error Resume Next
    s = c(5)
    s = "after:" & Err.Number
    Err.Clear
    On Error GoTo Handler
    Err.Raise 1004, , "boom"
    s = s & "not reached"
Done:
    T = s
    Exit Function
Handler:
    s = s & "|caught " & Err.Number & " " & Err.Description
    Resume Done
End Function"#;
    assert_eq!(text(run(src, "T")), "after:5|caught 1004 boom");
}

#[test]
fn resume_next_in_handler_continues_after_the_error() {
    let src = r#"
Function T() As String
    Dim s As String
    On Error GoTo H
    s = "a"
    Err.Raise 5
    s = s & "b"
    T = s
    Exit Function
H:
    s = s & "h"
    Resume Next
End Function"#;
    assert_eq!(text(run(src, "T")), "ahb");
}

#[test]
fn unhandled_error_reports_number() {
    let (r, _) =
        run_with(&[("Module1", ModuleType::Standard, "Sub T()\n Dim c As New Collection\n x = c(1)\nEnd Sub")], "T");
    assert_eq!(r.unwrap_err().number, 5); // host: Collection index out of range is error 5
}

#[test]
fn dictionary_via_create_object() {
    let src = r#"
Function T() As String
    Dim d As Object, k
    Set d = CreateObject("Scripting.Dictionary")
    d("b") = 2
    d.Add "a", 1
    For Each k In d.Keys
        T = T & k & "=" & d(k) & ";"
    Next
    T = T & d.Exists("a") & d.Count
End Function"#;
    assert_eq!(text(run(src, "T")), "b=2;a=1;True2");
}

#[test]
fn class_modules_with_properties() {
    let class = r#"
Private pName As String
Public Hits As Long
Private Sub Class_Initialize()
    Hits = 100
End Sub
Public Property Get Name() As String
    Name = pName
End Property
Public Property Let Name(v As String)
    pName = UCaseLike(v)
End Property
Private Function UCaseLike(v As String) As String
    UCaseLike = v & "!"
End Function
Public Function Greet() As String
    Hits = Hits + 1
    Greet = "hi " & Name
End Function"#;
    let main = r#"
Function T() As String
    Dim p As Person
    Set p = New Person
    p.Name = "bob"
    T = p.Greet() & " " & p.Hits & " " & TypeName(p)
End Function"#;
    let (r, _) = run_with(&[("Person", ModuleType::Class, class), ("Module1", ModuleType::Standard, main)], "T");
    assert_eq!(text(r.unwrap()), "hi bob! 101 Person");
}

#[test]
fn module_variables_constants_enums_and_statics() {
    let src = r#"
Public Const RATE As Double = 2
Private counter As Long
Enum Color
    Red = 1
    Green
    Blue = 10
End Enum
Function NextId() As Long
    Static n As Long
    n = n + 1
    NextId = n
End Function
Function T() As String
    counter = counter + 5
    NextId
    NextId
    T = counter * RATE & "," & Green & "," & Color.Blue & "," & NextId()
End Function"#;
    assert_eq!(text(run(src, "T")), "10,2,10,3");
}

#[test]
fn debug_print_and_msgbox_reach_the_host() {
    let (r, host) = run_with(
        &[("Module1", ModuleType::Standard, "Sub T()\n Debug.Print \"a\"; 1\n MsgBox \"hello\"\nEnd Sub")],
        "T",
    );
    r.unwrap();
    assert_eq!(host.printed, vec!["a1".to_string()]);
    assert_eq!(host.messages, vec!["hello".to_string()]);
}

#[test]
fn endless_loops_are_stopped() {
    let mut host = NoHost::default();
    let mut it = Interpreter::new(&mut host);
    it.limits.max_steps = 10_000;
    it.add_module("M", ModuleType::Standard, "Sub T()\n Do\n Loop\nEnd Sub").unwrap();
    assert_eq!(it.run("T", vec![]).unwrap_err().number, -1);
}

#[test]
fn native_calls_and_file_io_are_refused() {
    let src = "#If VBA7 Then\nPrivate Declare PtrSafe Function GetTickCount Lib \"kernel32\" () As Long\n#End If\nFunction A()\n A = GetTickCount()\nEnd Function\nSub B()\n Kill \"scratch.txt\"\nEnd Sub\nSub C()\n Open \"x\" For Output As #1\nEnd Sub\nFunction D()\n Set D = CreateObject(\"WScript.Shell\")\nEnd Function";
    for entry in ["A", "B", "C", "D"] {
        let (r, _) = run_with(&[("M", ModuleType::Standard, src)], entry);
        let e = r.unwrap_err();
        assert!(e.number == 70 || e.number == 429, "{entry}: {e}");
    }
}

/// A host that raises `Worksheet_Change` on Sheet1 whenever `Cell = v` is assigned.
#[derive(Default)]
struct EventHost {
    cell: Value,
    queued: Vec<macroost_lang::HostEvent>,
    log: Vec<String>,
}

impl macroost_lang::Host for EventHost {
    fn global(&mut self, name: &str, _args: &macroost_lang::Args) -> Option<macroost_lang::Result<Value>> {
        name.eq_ignore_ascii_case("cell").then(|| Ok(self.cell.clone()))
    }
    fn global_assign(
        &mut self,
        name: &str,
        _args: &macroost_lang::Args,
        value: Value,
    ) -> Option<macroost_lang::Result<()>> {
        if !name.eq_ignore_ascii_case("cell") {
            return None;
        }
        self.cell = value.clone();
        self.queued.push(macroost_lang::HostEvent {
            module: "Sheet1".into(),
            procedure: "Worksheet_Change".into(),
            args: vec![value],
        });
        Some(Ok(()))
    }
    fn debug_print(&mut self, text: &str) {
        self.log.push(text.to_owned());
    }
    fn take_events(&mut self) -> Vec<macroost_lang::HostEvent> {
        std::mem::take(&mut self.queued)
    }
}

#[test]
fn events_run_document_module_handlers_and_nesting_is_capped() {
    let sheet = "Private Sub Worksheet_Change(ByVal Target As Variant)\n Debug.Print \"changed \" & Target\n Cell = Target + 1\nEnd Sub";
    let main = "Sub T()\n Cell = 1\nEnd Sub";
    let mut host = EventHost::default();
    {
        let mut it = Interpreter::new(&mut host);
        it.add_module("Sheet1", ModuleType::Document, sheet).unwrap();
        it.add_module("Module1", ModuleType::Standard, main).unwrap();
        it.run("T", vec![]).unwrap();
        assert!(it.fire("ThisWorkbook", "Workbook_Open", vec![]).unwrap().is_none());
    }
    // The handler re-triggers itself; the chain stops after a bounded number of levels.
    assert!(host.log.len() > 3 && host.log.len() <= 17, "{}", host.log.len());
    assert_eq!(host.log[0], "changed 1");
}
