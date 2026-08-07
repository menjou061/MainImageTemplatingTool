# v1.3.0

- Four channel profiles now use canonical-first dynamic field matching.
- The canonical product image binding is `商品图` in Excel and `!商品图` in PSD.
- Historical image headers remain compatible as aliases.
- Dynamic text fields use matching `@字段名` PSD layers; unmatched non-empty fields are recorded as a non-blocking warning.
- Acceptance plan includes additive text and image bindings: an Excel `动态测试字段` must reach `@动态测试字段`, and an Excel `动态测试图片` must replace `!动态测试图片`; the three-gift hygiene layout must pass regression at the same time.
- Versioned Excel and PSD standard templates are supplied under `标准模板归档/v1.3_动态字段标准/`.
- The Windows launcher hides its CMD console by default; set `L0_SHOW_CONSOLE=1` only when collecting startup diagnostics.
