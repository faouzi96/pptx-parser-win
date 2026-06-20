# Publishing Plan for `@branding-review/pptx-parser`

Repository: https://github.com/faouzi96/pptx-parser-win

## 1. Prepare the repository

- Ensure `package.json` includes:
  - `main: "dist/index.js"`
  - `types: "dist/index.d.ts"`
  - `bin: { "pptx-parser": "dist/index.js" }`
  - `files: ["dist", "README.md"]`
  - `publishConfig.access: "public"`
  - `prepare` and `prepublishOnly` scripts to build before publishing
- Ensure `index.ts` re-exports `parsePptx` and any public type definitions.
- Make sure the package is not private.
- Add a meaningful `README.md` and `keywords`, `repository`, `bugs`, `homepage` fields.
- Confirm the repo origin is `https://github.com/faouzi96/pptx-parser-win`.

## 2. Local publish workflow

1. Build the package:
   ```bash
   npm run build
   ```
2. Inspect the output:
   ```bash
   ls dist
   ```
3. Log in to npm if needed:
   ```bash
   npm login
   ```
4. Publish the package:
   ```bash
   npm publish --access public
   ```
5. For a new release version:
   ```bash
   npm version patch
   git push --follow-tags
   ```

## 3. GitHub Actions CI/CD publish workflow

Create a workflow at `.github/workflows/publish.yml` that:

- runs on `push` to tags like `v*`
- builds the package
- validates the package
- publishes using `NPM_TOKEN`

### Required GitHub secret

- `NPM_TOKEN` containing an npm access token with publish rights.

### Recommended workflow triggers

- `push` on tags: `refs/tags/v*`
- `workflow_dispatch` for manual runs

## 4. Manual `gh` workflow generation

To create the workflow file manually using `gh`:

```bash
mkdir -p .github/workflows
cat > .github/workflows/publish.yml <<'EOF'
# content here
EOF
```

To run the published workflow:

```bash
gh workflow run publish.yml
```

## 5. Notes for the package user experience

- The same root entrypoint supports both imports and CLI execution.
- Consumers can do:
  ```ts
  import { parsePptx } from '@branding-review/pptx-parser';
  ```
- CLI users can do:
  ```bash
  npx pptx-parser file.pptx ./output
  ```

## 6. Security and publishing reminders

- Keep `NPM_TOKEN` secret in GitHub repository settings.
- Only publish from a trusted machine or CI environment.
- Use semantic versioning for releases.
