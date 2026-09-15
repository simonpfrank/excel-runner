# Release And Repository Migration Completion Plan

Status: In discovery. This document records the release work, migration work, and additions agreed
for the current completion cycle. Each item will be refined with verified repository evidence.

## Release Goal

Produce a deployable Excel Runner package that can be installed from its repository, invoked from
the command line, and prepared for transfer to the destination repository without losing its
workflow, testing, documentation, or quality guarantees.

## Current-Cycle Scope

- Make the project installable and runnable as a Python package from a clean environment.
- Prepare the agreed CED/Unify source deployment: a flat destination root with `cli.py` and its
  sibling runtime modules, invoked through `python cli.py workflow.yaml`.
- Define and validate a minimal distributable file set for the destination repository.
- Suppress the known openpyxl data-validation warning at the appropriate boundary without hiding
  unrelated warnings.
- Reach a completed test suite with 100% tests passing, including a reliable treatment of
  live-Excel/xlwings tests.
- Install Pylint, declare it in project dependency/configuration metadata, and resolve all agreed
  quality-check issues.
- Review the legacy implementation for reusable actions not yet available in Excel Runner.
- Decide whether a numeric parsing action is a release requirement, a deliberately deferred
  backlog item, or unnecessary because existing YAML value handling is sufficient.
- Bring the README, specification, YAML authoring skill, progress tracker, and test summary into
  agreement with current behavior.
- Add or amend tests, documentation, and release metadata for every functional addition made in
  this cycle.

## Documentation And README Work

- Reorganize the README so installation and first command-line run appear near the top.
- Present two desktop installation paths: ordinary package use and local development; specify
  when `requirements.txt` applies.
- Put a prominent safety warning directly beside the first-run instructions: never first execute a
  workflow against shared-folder paths. Collect local copies of every input/template workbook,
  configure the workflow to use those local files, and validate the run locally. Only after the
  workflow is confirmed correct should paths be changed to approved shared-folder locations.
- Document CED's primary invocation before other usage material: `python cli.py <workflow.yaml>`
  from the flat CED deployment root. Document package interfaces and repository development as
  compatibility/contributor alternatives.
- Explain all command-line arguments and environment overrides following the invocation examples.
- Move the README change log below installation and run instructions.
- State that users do not choose between openpyxl and xlwings: Excel Runner selects the safe and
  eligible backend, while retaining workflow-level consistency and transaction protection.
- Add an early file-handling section that explains source files, templates, destinations, scratch
  staging, commit behavior, and exactly which file paths can change for successful and failed
  runs.
- Describe all currently supported actions, including additions not yet reflected in the README.
- Correct `read_range` documentation: `sheet` is required for an A1 `range`, but optional when
  `range` is a workbook-level defined name because the name supplies its destination sheet. Add
  a named-range example with no `sheet:` field.
- Rename the preflight section to `Preflight Checks`.
- Place the audit-log introduction immediately after workbook lifecycle documentation.
- Rename the development section to `Development And Contributing` and add concise project
  contribution expectations: TDD, unit and integration tests, quality checks, and updates to the
  README, progress tracker, test summary, YAML skill, and specification.
- Add a `Running In Unify` section with `TBC` content until operational instructions are supplied.

## Quality And Test Completion

- Maintain 100% automated-test success with no skipped tests in the supported Windows-and-Excel
  release environment; record the final command, count, duration, and coverage result.
- Preserve regression coverage for the corrected audit-record and defined-name write tests.
- Keep live-Excel coverage enabled in the full suite. The suite takes several minutes because it
  starts real Excel instances; test runners must allow that duration rather than treating it as a
  hang.
- Ensure Pylint is installed and declared in the project, then address Pylint findings.
- Correct tool scope/configuration so Vulture, Pyright, and Mypy check project code rather than
  the virtual environment; resolve or explicitly configure every remaining agreed finding.
- Preserve the maximum permitted Radon complexity grade and add tests for any refactoring needed
  to meet other checks.

## Package And Unify Deployment Decision

Agreed initial deployment model: flat-root source-module deployment. Transfer `cli.py` and its
runtime-module siblings to the CED root and invoke the runner directly:

```powershell
python cli.py workflow.yaml
```

This intentionally does not depend on building or publishing a package artifact in this
repository, and no CI pipeline is required for the initial release. Runtime dependencies still
need a defined installation step in the target environment.

Decisions still required before implementation:

1. Exact Unify Python runtime, dependency installation capability, working directory, workbook
  and Excel availability, and approved source-file distribution method.
2. Whether Unify can install from `requirements.txt`, needs a local wheelhouse, or has another
  approved dependency mechanism.
3. Whether xlwings/Excel automation is available in Unify, and the required behavior if it is not.
4. Versioning, release tag, destination repository ownership, and the migration procedure.

The destination layout must keep direct `cli.py` execution reliable from the CED root. The source
transfer must include `cli.py`, `core.py`, `runner.py`, `engine.py`, `actions.py`, `backends.py`,
`requirements.txt`, and the operational YAML/workbook assets required by the specific Unify
workflow, while excluding tests, local fixtures, run outputs, and temporary reports.

## Legacy Action Review

- Inventory legacy actions and compare them to the current action registry and documented YAML
  vocabulary.
- Classify each legacy action as already implemented, a candidate for this release, unsuitable,
  or backlog.
- For candidates, define user value, YAML contract, backend support, validation, tests, and
  documentation before adding any implementation.

## Numeric Parsing Decision

- Assess whether YAML scalar typing and existing template resolution cover the actual use cases.
- Compare a possible `parse_number` action with the existing date parsing action: accepted input,
  locale/thousands/decimal rules, output type, validation, and error contract.
- Do not add a generic conversion action without a concrete workflow requirement and unambiguous
  parsing rules. Record the resulting decision in this document and the backlog if deferred.

## Release Readiness Additions To Evaluate

- Source-module deployment and clean-environment execution verification using `cli.py`.
- Exact runtime file manifest for the destination repository and Unify workflow.
- Version source, semantic-version policy, release notes, and Git tag procedure.
- Dependency pinning/constraints, supported Python versions, and optional xlwings dependency
  policy.
- A future standard-package/wheel option, if deployment needs eventually justify it.
- License, security/dependency audit, supported-platform statement, and error/exit-code contract.
- Reproducible examples and a clean-worktree release checklist.

## Investigation Record

### Confirmed Current State

- `pyproject.toml` already defines an installable `setuptools` package named `excel-runner`,
  version `0.1.0`, runtime dependencies, and a console entry point:
  `excel-runner = "excel_runner.cli:main"`.
- The installed console command, `python -m excel_runner`, and direct
  `python excel_runner/cli.py` each display CLI help in the current environment. This proves
  only that the currently editable installation can import; it is not a clean-install or wheel
  validation.
- Internal modules use bare imports, such as `from core import ExcelRunnerError`. The package
  currently works by adding its own directory to `sys.path` in `excel_runner/__init__.py`. This
  is fragile, obscures the import boundary, and causes a material portion of strict Mypy errors.
- `requirements.txt` contains only the four runtime dependencies and is UTF-16 encoded.
  `pyproject.toml` is the package dependency source and contains the developer dependency group.
- The current action registry contains 29 actions. README status says 28, while the test summary
  says 20 and 285 total tests. The latest collection found 601 tests.
- The existing YAML skill correctly describes template-first staging, but it contains stale copy
  semantics and must be verified against every registered action before release.
- `Progress_Tracker.md` was last updated for a 2026-09-02 status snapshot and does not describe
  the current template-first lifecycle, current test counts, recent actions, or current quality
  failures.
- `Specification.md`, `PRD.md`, `README.md`, and `test_summary.md` all require a full action and
  behavior reconciliation rather than isolated wording edits.

### Completed Additions In This Cycle

- `template:` now supplies the workbook source for every run, even when its destination `file:`
  already exists. The template is staged into scratch and successful output is committed to
  `file:`.
- `create_if_missing` has no effect when `template:` is present; with no template it still
  creates a blank workbook only when `file:` does not exist.
- Added regression coverage for a template overriding an existing output and validated the
  lifecycle-focused suite: 93 passed.
- Updated the README, workflow YAML skill, and full-showcase README for the template lifecycle.
- Added `parse_date` and its YAML documentation. The release documentation review must confirm
  that every action addition in the current cycle is documented consistently everywhere.
- Ran a quality-gate investigation and saved its evidence in `docs/temp_quality report.md`.

### Agreed Deployment Approach

#### Flat-root source-module deployment

Distribute these files at the CED root: `cli.py`, `core.py`, `runner.py`, `engine.py`,
`actions.py`, `backends.py`, and `requirements.txt`. Then invoke:

```powershell
python cli.py workflow.yaml
```

Advantages: it matches the required Unify invocation, does not require a build step in this
repository, and makes the deployed runtime files explicit.

Constraints to address: dependency installation remains required; imports must continue to work
when `cli.py` is invoked directly from the flat CED root; and the transferred file set
must be explicitly versioned and documented.

#### Deferred Package Option

`pyproject.toml` already supports a standard package, console script, and wheel build. Retain
that metadata for later use, but do not make wheel build, package installation, or CI pipeline
work a blocker for this release cycle.

### Source-Module Release Implementation Plan

1. Confirm the target release version, supported Python version, destination repository, and
  whether migration preserves Git history or copies a selected source snapshot.
2. Define a minimal flat-root source deployment manifest: `cli.py`, `core.py`, `runner.py`,
  `engine.py`, `actions.py`, `backends.py`, `requirements.txt`, the target workflow YAML file(s),
  and only the runtime assets each workflow needs. Explicitly exclude tests, documentation drafts,
  `temp` fixtures, run outputs, caches, and `.venv`.
3. Make direct execution robust from the approved destination repository working directory.
  Preserve or replace the current import handling only after a test proves
  `python cli.py workflow.yaml` works from the flat CED root.
4. Define runtime versus developer dependencies. Document `pip install -r requirements.txt` for
  Unify and desktop source deployment. Convert `requirements.txt` to UTF-8 and decide how it is
  kept synchronized with runtime dependencies in `pyproject.toml`.
5. Decide whether the target workflow needs `xlwings` and Excel. If it does, make the Windows
  Excel prerequisite explicit; if it does not, avoid deploying or invoking the COM-only paths.
6. Run a clean-environment source deployment smoke test using the manifest: install runtime
  dependencies, copy the manifest to a temporary destination, run `cli.py --help`, a dry-run
  workflow, and a normal non-COM workflow.
7. Prepare release notes, migration instructions, version tag, license/security review, and a
  clean release checklist before transferring to the destination repository.
8. Revisit a standard wheel/console-script release and CI pipeline only when deployment volume or
  operational controls make their maintenance cost worthwhile.

### Running In Unify

TBC

Before implementation, provide: Unify operating system and Python version; whether it can install
wheels/dependencies; package-index policy; command execution and working-directory rules;
workbook and network-storage access; Excel/xlwings availability; and artifact/repository access.

### Warning Suppression Plan

The requested openpyxl data-validation warning is not emitted by project code. Workbook loading
is centralized in a small number of locations in `backends.py` and preflight existence validation
in `engine.py`.

1. Reproduce the warning against the identified workbook and capture its exact message and warning
  category.
2. Add a tightly scoped `warnings.catch_warnings()` filter around only the relevant
  `openpyxl.load_workbook()` boundary. Match the exact message/category rather than suppressing
  all `UserWarning` values or all openpyxl warnings.
3. Add a regression test proving that the known data-validation warning is suppressed while an
  unrelated openpyxl warning remains visible.
4. Document the rationale and source-workbook compatibility limitation in the specification and
  changelog.

### Quality And Test Remediation Plan

#### Completed Test Remediation

1. `TestStop.test_stopped_steps_still_get_an_audit_record` was stale: `audit.jsonl` deliberately
   includes both run-level `event` records and per-step `step_id` records. The assertion now
   filters step records, matching the established audit-file contract. Focused audit tests pass.
2. The two defined-name write tests were stale: `write_cell` and `write_range` accept the target
   before the optional sheet. Their positional calls now match the public action signatures and
   correctly verify named-target behavior. Focused tests pass.
3. The former full-suite "hang" was an external execution wrapper's two-minute capture limit,
   not a blocked test. The real Excel backend contract alone takes 86 seconds and the ordered
   backend suite takes 157 seconds. Do not add a skip gate for supported Windows Excel runs.
4. Full no-skip validation completed on 2026-09-15:

   ```powershell
   .venv\Scripts\pytest -q -rs
   ```

   Result: `601 passed in 301.05s (0:05:01)`, with zero skipped tests.

#### Ongoing Full-Suite Requirement

Run the full suite in the supported Windows-and-Excel environment after each substantial release
slice. Release acceptance is 100% selected tests passing with zero skips. Configure external test
runners to permit at least the observed five-minute execution time.

#### Quality Tooling

1. Add `pylint` to `project.optional-dependencies.dev`, install it into `.venv`, and add a
  deliberate `[tool.pylint.*]` configuration or command scope. Run it only over repository
  source/tests, never `.venv`.
2. Make Vulture scope explicit and apply its whitelist only to confirmed framework/dynamic uses.
  Its current `vulture .` command scans virtualenv packages and exits `3`, obscuring project
  findings. Resolve or whitelist the five known repository/test findings with justification.
3. Add a `pyrightconfig.json` or equivalent configuration pointing to `.venv` and excluding
  `.venv`, run artifacts, fixtures, and generated outputs. Then resolve actual project errors.
4. Configure Mypy to check the installed package structure rather than using `.` recursively.
  The package-relative import refactor should remove the current import-not-found cluster;
  type the action decorator and correct the remaining source/test annotations.
5. Keep Radon no worse than grade C. Its current grade-C results are acceptable, but validate that
  any quality refactor does not create grade D/E paths.
6. Make the final quality gate a single documented, repeatable command or CI job, with explicit
  tool scopes and outputs. The release cannot claim a clean gate until all configured checks pass.

### README Rewrite Plan

Required order:

1. Excel Runner introduction: declarative workflow automation; no user decision between openpyxl
  and xlwings; automatic backend eligibility/selection; scratch staging and atomic commit;
  preflight checks; structured audit/run artifacts; and automatic close/cleanup.
2. Installation and first run: side-by-side desktop source deployment and local-development
  paths, both using `pip` commands. Document `pip install -r requirements.txt` as the runtime
  dependency installation for desktop and Unify source-module deployment. Explain that
  `pip install -e ".[dev]"` is for contributors working in this repository. Place the following
  first-run safety warning prominently in this section:

  > Do not run a new or changed workflow against shared-folder paths first. Collect local copies
  > of the required workbooks, run and validate the workflow locally, then update paths to
  > approved shared folders only after the results are confirmed.

3. Command line: direct `python cli.py workflow.yaml` first for CED source deployment, with all
  runtime modules placed at the CED root. List `python -m excel_runner` and `excel-runner` as
  optional, future package-installation alternatives rather than the production path for this
  cycle. Use actual Windows examples and explain all arguments: workflow, repeatable `--env KEY=VALUE`,
  `--working-dir`, `--logging-level`,
  `--no-logfile`, `--check-existence`, and `--dry-run`.
  State that both execution modes are covered by automated smoke tests: direct root-level
  `python cli.py workflow.yaml` execution for the CED source bundle, and package-style
  `python -m excel_runner workflow.yaml` / `excel-runner workflow.yaml` execution for the
  repository package interface. The CED mode is the initial deployment target; package-mode
  coverage protects the existing public interface for future deployment use.
4. Change log immediately after installation and run instructions.
5. Quick-start workflow and YAML overview, followed by a callout: workflows can be authored
  manually or by asking GitHub Copilot to use the `excel-runner-yaml` skill. Link the skill and
  state that it is the authoritative syntax reference.
6. `How Files Are Handled`: place this as high as practical after workflow basics and before the
  action reference. Include a concise source-to-scratch-to-destination diagram and a table for
  these cases:

  | Workflow configuration and outcome | Starting source | Can change | Must not change |
  | --- | --- | --- | --- |
  | `template:` declared, successful run | template copied to scratch every run | destination `file:` after commit | template and original shared/local source |
  | `template:` declared, failed run | template copied to scratch | scratch/run artifacts only | destination `file:` and template |
  | no template, existing `file:`, successful run | `file:` copied to scratch | that destination `file:` after commit | original only until commit |
  | no template, existing `file:`, failed run | `file:` copied to scratch | scratch/run artifacts only | destination `file:` |
  | no template, missing `file:`, `create_if_missing: true` | new blank scratch workbook | destination `file:` after commit | unrelated files |

  Explain that scratch is an isolated working copy in the run's working directory. Actions open
  and modify scratch, not declared source/destination files directly. A successful run closes,
  saves, and atomically commits changed scratch workbooks to their destinations; a failed run
  does not commit them. Explain that `template:` is a good fit for repeatable output generation
  because it resets the starting content every run, preventing an old output from becoming the
  next run's hidden input. State any exceptions to these guarantees explicitly, including
  external Excel links or actions that intentionally touch external resources.
7. `Save Blockers`: place this subsection immediately after scratch/lifecycle behavior and before
   the audit-log introduction. Explain the safety problem in user terms: openpyxl cannot safely
   save a workbook with outbound external workbook links because it can damage the link metadata,
   even when the workflow did not intend to change those links. Explain the current behavior:

   - Excel Runner inspects the staged OOXML workbook before opening it for work.
   - The only verified save blocker today is one or more outbound external links; inbound links
     from other workbooks do not block saving this workbook.
   - Read-only access remains on the faster openpyxl backend because it cannot save or corrupt the
     file.
   - Before the first write, a blocker-bearing read-write workbook is automatically promoted to
     Excel/xlwings. It remains on that backend for the run so a later openpyxl save cannot damage
     links.
   - A defensive guard refuses a save if an internal routing error ever leaves a dirty,
     blocker-bearing workbook on openpyxl. The workflow stops instead of writing a corrupted file.
   - This requires a usable local Excel/xlwings environment for any workflow that writes a
     blocker-bearing workbook. If that capability is unavailable, the run must report the clear
     prerequisite/error rather than fall back to an unsafe save.

   Include a brief example of an output workbook containing formulas linked to another workbook.
   State that save blockers are an intentionally small, evidence-based list: additional blockers
   are added only after they are empirically shown unsafe for openpyxl round-tripping. Mention
   that the audit log records detected blockers and backend switching.
8. Audit log introduction immediately after `Save Blockers`.
9. Complete action reference generated/reconciled against all 29 registered actions. Correct
  `read_range`: `sheet` is mandatory for A1 notation but optional for a workbook-level defined
  name; add a named-range example without `sheet:`. Include all current action arguments,
  backend implications, output shapes, and known limitations.
10. `Preflight Checks`, then library usage, `Running In Unify` with `TBC`, current limitations,
  and `Development And Contributing`.
11. Development/contributing bullets: TDD, unit tests, real integration tests, required quality
   checks, and required maintenance of README, progress tracker, test summary, YAML skill, and
   specification for every behavior change.

### Documentation Reconciliation Plan

1. Treat source plus tests as behavioral truth. Build an action/parameter/output matrix from the
  registry and action signatures before editing prose.
2. Update the YAML skill from that matrix, including current copy backend semantics, defined-name
  support, template-first lifecycle, preflight commands, action output keys, and all 29 actions.
3. Update the specification's models, lifecycle diagrams/text, execution/backend logic, CLI
  contract, validation tiers, and action catalog to current behavior.
4. Update the PRD only where it describes delivered behavior, release acceptance criteria, or
  intentional backlog. Preserve product decisions that still need user agreement.
5. Replace the stale progress tracker snapshot with a concise current state, completed work,
  active release work, owner decisions, quality status, and backlog.
6. Regenerate the test summary from the final suite rather than hand-maintaining stale counts.

### Legacy Review Decision

The useful legacy behavior is already represented by current actions or workflow composition:
column copying, table-cell updates, period replacement, workbook/template lifecycle, FAC/CSV
reading, named-range writes, placeholder substitution, and recalculation. Do not reintroduce the
legacy dispatcher policy that logs invalid steps and continues: Excel Runner's preflight plus
fail-fast execution is safer for audited financial workflows.

Two items need explicit release decisions:

- A row-wide period update is potentially useful, but no current workflow requires it. Define the
  desired row-selection contract and prove that `replace_table_text` plus existing table actions
  cannot express it before adding a new action.
- The reference-workbook comparator is useful release validation tooling, but it is not a general
  runtime action. Decide whether it becomes a documented test utility or stays fixture-specific.

### Numeric Parsing Decision

Recommendation: backlog `parse_number`; do not include it in this release without a real workflow
need. YAML native scalars and whole-template native typing already preserve legitimate numeric
values, while `read_text_file` intentionally preserves identifiers such as `00123` and periods
such as `202606` as text. A number parser must first define decimal separators, thousands
separators, signs, whitespace, null values, locale, and the distinction between identifiers and
quantities. Adding it solely because `parse_date` exists would create an unsafe conversion API.

### Release Exit Criteria

- Source-module deployment manifest and clean-environment `cli.py` smoke tests pass.
- Unify deployment model and operational instructions are agreed and documented.
- Every selected test passes with no skipped tests in the supported Windows-and-Excel environment.
- The full test runner permits the observed five-minute live-Excel execution time.
- All configured quality checks pass from defined project scope.
- Pylint is installed and declared in `pyproject.toml`.
- The data-validation warning is narrowly suppressed and regression-tested.
- README, PRD, specification, YAML skill, progress tracker, and test summary agree with source
  behavior and test evidence.
- New action decisions are implemented and documented, or explicitly deferred with rationale.
- Release notes, version, source deployment manifest, destination-repository migration, and
  rollback procedure are reviewed before transfer.