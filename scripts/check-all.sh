#!/usr/bin/env bash
# Full gate: format, lints, tests, and a compile check for every platform we ship on.
set -euo pipefail
cd "$(dirname "$0")/.."
export CARGO_TARGET_DIR="${CARGO_TARGET_DIR:-target}"
cargo fmt --all -- --check
cargo clippy -j 2 --workspace --all-targets -q -- -D warnings
cargo test -j 2 --workspace -q
for t in x86_64-unknown-linux-gnu aarch64-apple-darwin x86_64-apple-darwin wasm32-unknown-unknown; do
  cargo check -j 2 -q --workspace --target "$t" && echo "ok $t"
done
