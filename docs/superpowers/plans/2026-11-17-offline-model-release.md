# Offline Model Release Implementation Plan

**Goal:** Ship a Windows ZIP that contains the application and its complete multilingual semantic model, and publish it as a GitHub Release asset without committing model weights to Git.

**Architecture:** The matcher loads the bundled model from `models/paraphrase-multilingual-MiniLM-L12-v2` with Hugging Face local-only mode, then may use an existing local cache; it never falls back to a network download. `run.bat` validates imports and package consistency on every launch and repairs or stops before the GUI starts. The ZIP contains the model and source while Git ignores the model directory and ZIP.

**Tech stack:** Python 3.11/3.12, Sentence Transformers, unittest, Windows batch, ZIP, GitHub.

## Requirements and constraints

- Preserve the GUI, matcher, review flow, Excel report, and existing tests.
- Include actual weights, tokenizer, Sentence Transformers modules, configs, and a license copy in the offline ZIP.
- Keep the model and ZIP out of Git; commit and push code only.
- Verify a real model load and semantic match with Hugging Face offline flags enabled before publishing.
- Publish `ExcelTracker_Offline.zip` as a Release attachment and verify its size and link.

## Implementation sequence

1. Add tests for bundled-model precedence and strictly local-only model construction; update the matcher and user-facing errors.
2. Update `run.bat` so incomplete environments are repaired and rechecked, otherwise launch stops.
3. Update README for the offline release and model provenance/license; ignore bundled model data and the generated ZIP in Git.
4. Copy the verified cache snapshot, include Apache-2.0 license text, build the ZIP, inspect its entries and sizes.
5. Run the complete unit suite and a real semantic match in a fresh process with Hugging Face offline mode enabled.
6. Commit/push source changes without model weights, create the GitHub Release, upload the ZIP, then verify the Release asset.

## Review focus

- Missing model or dependency files: matcher must fail with actionable local-only guidance; `run.bat` must not start the GUI.
- Present bundled model: it must load from the project path even when Hugging Face network access is disabled.
- Archive construction: model safetensors, tokenizer files, module configuration, source, README, requirements, and launcher must all be present.
- Git changes: model weights and the generated ZIP must remain untracked/ignored while source changes are committed.

