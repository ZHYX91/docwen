# Display and task languages

`gui.language.locale` controls DocWen's interface text and layout direction. A consumer's interface language controls that consumer's own text and layout. Neither value is an implicit conversion option. Resource names may be translated for display; canonical resource IDs, capability IDs, schemas and error codes remain language-independent.

`conversion.output_locale`, exposed as **Task output language** in General settings, selects generated labels, managed DOCX styles, localized placeholder processing and DOCX proofreading comment text. It does not translate source text, switch a selected template, change proofreading rule dictionaries or override an explicit OCR model. New profiles use `zh_CN`. A supported explicit task `options.locale` takes precedence. CLI `--lang` is an explicit locale request and is forwarded only to routes that declare that option; initializing a CLI translator alone does not set task options. Consumers that omit the option use Core's independent output setting.

When an old user profile has an explicitly saved display locale but no output locale, loading it freezes that saved value once into its sparse `conversion.toml` override. Existing output overrides win. Later display-language saves preserve the independent output value. No template state or proofreading dictionary is rewritten. Explicit task options remain authoritative. Configuration imports and resets must preserve this distinction.

`ocr.language` remains a separate recognition setting. Explicit task `ocr_language` takes precedence. Its `auto` value means **Follow task output language**, not image-language detection or following interface language. A fixed task/output language therefore chooses the same model after a display-language change. Arabic output does not add an Arabic OCR model; the existing automatic mapping uses English for that locale.

OCR title overrides remain keyed by output locale. Generated YAML labels and OCR titles are resolved from the admitted request and its frozen configuration, rather than a process-wide UI translator. An explicit `yaml_key_labels` object remains an output override. Input content, source identities and machine field names are not translated.

Saved template IDs and manual choices remain authoritative. The one-time initial GUI template recommendation is retained; switching interface language after template state is initialized does not select another template. CLI/Machine requests without a template use the built-in document rather than that GUI recommendation.

Source regression, frozen-package verification and real Qt/consumer/Office acceptance are separate evidence. Interface-language differences alone must not make a fixed task unavailable or change its execution parameters.
