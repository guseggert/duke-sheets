# Bindings

- **[python/](./python/)** — PyO3 native extension
- **[nodejs/](./nodejs/)** — NAPI-RS native addon
- **[wasm/](./wasm/)** — wasm-bindgen for browser + Node.js

See [github.com/guseggert/duke-sheets](https://github.com/guseggert/duke-sheets).

Node.js, Python, and WASM expose worksheet form-control collection and mutation
APIs. The experimental C FFI does not yet expose worksheet objects.

## Building

All bindings are managed via [mise](https://mise.jdx.dev/):

```bash
mise run setup     # Install toolchain + deps
mise run build     # Build all bindings
mise run test      # Run all binding test suites
```
