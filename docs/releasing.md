# Releasing packages

The Word packages use the `@christophervr` npm scope and share a version. The public umbrella package is `@christophervr/docx-viewer`; installing it brings in the internal core, document, legacy, web-component and bindings packages.

## Current release scope

Publish `@christophervr/ole2` from the sibling repository first. The current release scope is that shared CFB and legacy Word binary package only. Do not tag or dispatch the Word release workflow yet; Word package publication and PowerPoint's planned `/embedded` migration wait until the shared dependency is available and consumer integration is ready.

For a registry owner preparing initial access, authenticate to npm with an account that can publish the `@christophervr` scope:

```sh
npm login --scope=@christophervr
```

Use the sibling repository's release instructions to publish `@christophervr/ole2@0.1.0`. The Word package release depends on that public registry package for legacy `.doc` support.

The Word release workflow supports npm trusted publishing through GitHub OIDC and an `NPM_TOKEN` fallback. Configure the npm trusted publisher separately for each Word package with this repository and `release.yml`; configure `ole2` against the sibling repository's publishing workflow. Keep each package's publisher entry aligned with its actual GitHub owner/repository and workflow path.

## Future Word release

When Word package publication and the PowerPoint migration are ready, all six Word packages are released at the same version. The release script builds, checks, packs, and publishes them in dependency order:

1. `@christophervr/docx-core` provides the shared DOCX model, parser, serializer, and embedded-DOCX API.
2. `@christophervr/docx-legacy` adapts legacy `.doc` documents using `@christophervr/ole2`.
3. `@christophervr/docx-document` detects DOCX and legacy DOC input and loads it through the two adapters.
4. `@christophervr/docx-web-component` provides the custom-element editor.
5. `@christophervr/docx-bindings` provides framework lifecycle adapters.
6. `@christophervr/docx-viewer` provides one-install root and framework subpath exports.

`ole2` is released separately before this sequence. Keep each Word manifest at the same version because internal Word package dependencies are version matched.

Run the release checks locally before creating the tag:

```sh
bun install --frozen-lockfile
bun run typecheck
bun run test
bun run build:packages
bun run pack:smoke
bun run release:check 0.1.0
```

After the Word release is explicitly scheduled and its registry dependency and trusted publisher configuration are ready, create and push the release tag:

```sh
git tag v0.1.0
git push origin v0.1.0
```

The tag starts the [release workflow](https://github.com/ChristopherVR/docx-viewer/blob/main/.github/workflows/release.yml), which validates the version and publishes the six Word packages. A manual workflow dispatch accepts the version as `0.1.0` or `v0.1.0`.
