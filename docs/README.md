# simple-excel docs site

The documentation for `@jetstreamapp/simple-excel`, built with [Docusaurus](https://docusaurus.io/) and published
at https://simple-excel.getjetstream.app. This is a separate npm project with its own `package-lock.json`. It
resolves the library from the repository's `dist/`, so build that first.

```bash
npm ci && npm run build      # in the repository root
cd docs && npm ci
npm start                    # local dev server
npm run build                # static site in build/, plus build/_headers and build/serve.json
npm run serve:csp            # serve build/ locally with the headers from build/_headers
```

`npm run build` runs `scripts/generate-csp.mjs`, which hashes every inline script in the build output into the
Content-Security-Policy written to `build/_headers` (applied by Cloudflare Pages and by `npm run serve:csp`) and
`build/serve.json` (applied by `npx serve build`). CI builds the site on pull requests that touch `docs/`; it does
not deploy it.
