# React + TypeScript + Vite

This template provides a minimal setup to get React working in Vite with HMR and some ESLint rules.

Currently, two official plugins are available:

- [@vitejs/plugin-react](https://github.com/vitejs/vite-plugin-react/blob/main/packages/plugin-react) uses [Oxc](https://oxc.rs)
- [@vitejs/plugin-react-swc](https://github.com/vitejs/vite-plugin-react/blob/main/packages/plugin-react-swc) uses [SWC](https://swc.rs/)

## React Compiler

The React Compiler is not enabled on this template because of its impact on dev & build performances. To add it, see [this documentation](https://react.dev/learn/react-compiler/installation).

## Expanding the ESLint configuration

If you are developing a production application, we recommend updating the configuration to enable type-aware lint rules:

```js
export default defineConfig([
  globalIgnores(['dist']),
  {
    files: ['**/*.{ts,tsx}'],
    extends: [
      // Other configs...

      // Remove tseslint.configs.recommended and replace with this
      tseslint.configs.recommendedTypeChecked,
      // Alternatively, use this for stricter rules
      tseslint.configs.strictTypeChecked,
      // Optionally, add this for stylistic rules
      tseslint.configs.stylisticTypeChecked,

      // Other configs...
    ],
    languageOptions: {
      parserOptions: {
        project: ['./tsconfig.node.json', './tsconfig.app.json'],
        tsconfigRootDir: import.meta.dirname,
      },
      // other options...
    },
  },
])
```

You can also install [eslint-plugin-react-x](https://github.com/Rel1cx/eslint-react/tree/main/packages/plugins/eslint-plugin-react-x) and [eslint-plugin-react-dom](https://github.com/Rel1cx/eslint-react/tree/main/packages/plugins/eslint-plugin-react-dom) for React-specific lint rules:

```js
// eslint.config.js
import reactX from 'eslint-plugin-react-x'
import reactDom from 'eslint-plugin-react-dom'

export default defineConfig([
  globalIgnores(['dist']),
  {
    files: ['**/*.{ts,tsx}'],
    extends: [
      // Other configs...
      // Enable lint rules for React
      reactX.configs['recommended-typescript'],
      // Enable lint rules for React DOM
      reactDom.configs.recommended,
    ],
    languageOptions: {
      parserOptions: {
        project: ['./tsconfig.node.json', './tsconfig.app.json'],
        tsconfigRootDir: import.meta.dirname,
      },
      // other options...
    },
  },
])
```

## Chart column transformations

Use `ColumnExpressionSelect` for chart dimension and filter column pickers, including pivot rows and columns. Date and datetime options display “Configure…” and open a configuration modal before committing a selection. The selected transformation appears as an editable badge; selecting the same column again reopens its configuration. Cancel leaves the existing selection unchanged, and selecting a different column starts with its raw value.

`columnExpressions.ts` owns transformation metadata, type eligibility, expression parsing, evaluation, and display labels. Persist expressions as a plain column path or `path:feature`, including reference paths such as `customer.created_at:weekofmonth`. Chart aggregation, pivot dimensions, filter options, and filter matching must use the same expression evaluator. Numeric value-column selectors use the same picker with transformations disabled.

Week of month uses Sunday-start calendar weeks numbered 1–6, matching the week-of-year convention. Week 1 contains the first day of the month; each Sunday starts the next week. Date features use the browser's local time, and hour is available only for datetime columns.

## Automatic Deploys On Push

This repo now includes a GitHub Actions workflow that builds and deploys to Cloudflare Pages on every push to `main`.

Workflow file:

- `.github/workflows/deploy-pages.yml`

Add these repository secrets in GitHub before the workflow runs:

- `CLOUDFLARE_API_TOKEN` (token with Cloudflare Pages edit permission)
- `CLOUDFLARE_ACCOUNT_ID` (your Cloudflare account ID)
- `CLOUDFLARE_PAGES_PROJECT` (your Pages project name)

After these secrets are set, every commit pushed to `main` will trigger a deploy automatically.
