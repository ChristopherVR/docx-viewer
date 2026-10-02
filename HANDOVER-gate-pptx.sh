#!/usr/bin/env bash
# Local trunk gate for pptx-viewer: typecheck + lint on changed files + all unit suites in parallel.
# usage: gate-pptx.sh <worktree>   (prints a one-line verdict per package, exit 1 on any failure)
set -u
cd "$1" || exit 2
out=$(mktemp -d)
fail=0

echo "== typecheck"
if ! bun run typecheck >"$out/typecheck.log" 2>&1 || grep -qE "error TS|Exited with code [1-9]" "$out/typecheck.log"; then
	grep -E "error TS|Exited with code [1-9]" "$out/typecheck.log" | head -5
	echo "typecheck FAILED"; fail=1
else echo "typecheck ok"; fi

echo "== lint (changed files vs origin/main)"
files=$(git diff --name-only origin/main...HEAD | grep -E '\.(ts|tsx|mjs|js)$' | while read -r f; do [ -f "$f" ] && echo "$f"; done)
if [ -n "$files" ]; then
	if ! echo "$files" | xargs bunx oxlint --deny-warnings >"$out/lint.log" 2>&1; then tail -8 "$out/lint.log"; echo "lint FAILED"; fail=1; else echo "lint ok"; fi
fi

echo "== unit suites (parallel)"
pids=()
for pkg in shared react vue svelte vanilla angular locales core; do
	[ -d "packages/$pkg" ] || continue
	( cd "packages/$pkg" && bunx vitest run >"$out/unit-$pkg.log" 2>&1; echo $? >"$out/unit-$pkg.code" ) &
	pids+=($!)
done
wait "${pids[@]}"
for pkg in shared react vue svelte vanilla angular locales core; do
	[ -f "$out/unit-$pkg.code" ] || continue
	line=$(sed 's/\x1b\[[0-9;]*m//g' "$out/unit-$pkg.log" | grep -E "^\s+Tests " | tail -1)
	if [ "$(cat "$out/unit-$pkg.code")" != "0" ]; then
		echo "unit $pkg FAILED: $line"; sed 's/\x1b\[[0-9;]*m//g' "$out/unit-$pkg.log" | grep -E "FAIL " | head -6; fail=1
	else echo "unit $pkg ok: $line"; fi
done
echo "GATE_EXIT=$fail"
exit $fail
