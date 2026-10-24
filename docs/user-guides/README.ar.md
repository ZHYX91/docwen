# DocWen

<p align="center">
  <img src="https://raw.githubusercontent.com/ZHYX91/docwen/main/assets/icon.svg" alt="DocWen logo" width="120">
</p>

[English](https://github.com/ZHYX91/docwen/blob/main/README.md) · [简体中文](https://github.com/ZHYX91/docwen/blob/main/docs/user-guides/README.zh-CN.md) · [繁體中文](https://github.com/ZHYX91/docwen/blob/main/docs/user-guides/README.zh-TW.md) · [Deutsch](https://github.com/ZHYX91/docwen/blob/main/docs/user-guides/README.de-DE.md) · [Français](https://github.com/ZHYX91/docwen/blob/main/docs/user-guides/README.fr-FR.md) · [Español](https://github.com/ZHYX91/docwen/blob/main/docs/user-guides/README.es-ES.md) · [Português](https://github.com/ZHYX91/docwen/blob/main/docs/user-guides/README.pt-BR.md) · [Русский](https://github.com/ZHYX91/docwen/blob/main/docs/user-guides/README.ru-RU.md) · [日本語](https://github.com/ZHYX91/docwen/blob/main/docs/user-guides/README.ja-JP.md) · [한국어](https://github.com/ZHYX91/docwen/blob/main/docs/user-guides/README.ko-KR.md) · [Tiếng Việt](https://github.com/ZHYX91/docwen/blob/main/docs/user-guides/README.vi-VN.md) · [العربية](https://github.com/ZHYX91/docwen/blob/main/docs/user-guides/README.ar.md)

يساعدك DocWen على تحويل المستندات والجداول والصور **على جهازك**، دون رفع الملفات إلى خدمة تحويل على الإنترنت. يمكنك استخدام الواجهة الرسومية للأعمال اليومية أو سطر الأوامر لتنفيذ عمليات متكررة.

تتحكم لغة الواجهة في نصوص التطبيق واتجاه عرضه فقط. تتحكم لغة إخراج المهمة في الإعدادات العامة بشكل مستقل في التسميات المنشأة وأنماط المستند وتعليقات التدقيق، ولا تترجم النص. لا يؤدي تغيير الواجهة إلى تغيير لغة الإخراج أو القالب المحفوظين. يتبع خيار OCR التلقائي لغة إخراج المهمة. يمكن اختيار إنشاء فقرة Word مستقلة لكل سطر Markdown عادي، وهذا الخيار معطل افتراضيا.

هذا دليل عربي مختصر للمهام الأساسية. تتوفر التفاصيل الكاملة عن الواجهة والخيارات والقوالب في [الدليل الإنجليزي](../../README.md) و[الدليل الصيني المبسط](README.zh-CN.md). وللتأكد من مسارات التحويل المتاحة راجع [جدول القدرات](../capabilities.md).

## ما الذي يمكن فعله؟

- تحويل ملفات Word إلى Markdown والعكس، مع الحفاظ على العناصر التي تدعمها طريقة التحويل المختارة.
- تصدير جداول Excel إلى Markdown وإنشاء جداول من ملفات Markdown.
- استخراج النص من PDF وXPS وOFD؛ والتعرف على النص داخل الصور محليًا عند توفر OCR.
- دمج ملفات PDF أو تقسيمها، وتحويل الصور إلى تنسيقات أخرى.
- اختيار قوالب Word أو Excel الخاصة بك، وفحص النصوص بقواميس وقواعد مخصّصة.

لا يدعم كل تنسيق جميع عمليات التحويل. يمكنك فحص الملف والعمليات المتاحة أولًا؛ راجع [جدول القدرات](../capabilities.md).

**قوالب Word:** لا يوجد قالب Word عربي مستقل ضمن البرنامج. عند إعداد الواجهة الرسومية بالعربية لأول مرة، يُختار القالب الإنجليزي العام. لا يؤدي تبديل اللغة إلى تغيير اختيار القالب المحفوظ. في سطر الأوامر، يُستخدم مستند A4 مدمج إذا لم تحدد قالبًا.

**التعرف على النص (OCR):** لا تتوفر حاليًا نماذج OCR للعربية. يتبع الخيار «تلقائي» **لغة إخراج المهمة**، وليس لغة الواجهة أو الاكتشاف التلقائي للغة الصورة. إذا كانت لغة إخراج المهمة عربية، يستخدم الربط الحالي نموذج الإنجليزية؛ ويمكن اختيار نموذج OCR آخر متاح يدويًا وفق نص الصورة.

## البدء بسرعة

### تنزيل البرنامج وتشغيله

نزّل الأرشيف المناسب من [صفحة الإصدارات](https://github.com/ZHYX91/docwen/releases). تتوفر حزم **Windows x64** و**Ubuntu 24.04 x64** و**Apple Silicon مع macOS 14+**، دون الحاجة إلى تثبيت Python. استخرج الأرشيف كاملًا: افتح `DocWen.exe` على Windows أو شغّل `./DocWen` على Ubuntu؛ على Mac انقل `DocWen.app` إلى التطبيقات واتبع [دليل الفتح الأول](macos.md). تتوفر أيضًا حزم CLI مستقلة لـ Ubuntu وmacOS. توزيعات Linux الأخرى وأجهزة Intel Mac ليست ضمن أهداف الحزم.

تختلف الملفات المتاحة حسب الإصدار. اختر أرشيفًا مدرجًا بالفعل لنظامك؛ قد لا تتضمن الإصدارات القديمة حزمة macOS.

### التشغيل من المصدر

للتطوير أو التثبيت من المصدر، ثبّت **Python 3.12** و**uv 0.12.0**، ثم نفّذ داخل المستودع:

```bash
git clone https://github.com/ZHYX91/docwen.git
cd docwen
uv sync --frozen --all-extras
uv run --frozen docwen-gui
uv run --frozen docwen --help
```

يستخدم `uv run --frozen` البيئة المتزامنة دون تفعيلها يدويًا. تستخدم اختبارات CI وبناء المنتج uv 0.12.0؛ وتقبل صيانة الاعتماديات إصدارات 0.12.x المتوافقة. لا يدعم المستودع `pip install -e` لأن pip لا يطبّق قاعدة استبعاد الاعتماديات المحددة فيه.

تتوفر حزم Apple Silicon macOS 14+ arm64 بتوقيع ad-hoc، دون Developer ID أو توثيق Apple. راجع [التثبيت والفتح الأول](macos.md). لا تتوفر حزمة Intel Mac.

### صور HEIC/HEIF

تستخدم الحزم محرّك الصور المضمّن فيها. افحص النسخة المثبّتة باستخدام `DocWenCLI doctor --json` (راجع مسار CLI على Mac في الدليل). تثبيت `pillow-heif` في Python النظام لا يغيّر التطبيق المجمّع. عند التشغيل من المصدر يثبّت `uv sync --frozen --all-extras` مكتبة `pillow-heif` بالفعل. إذا أبلغ محرّك macOS عن غياب مكتبة أصلية، نفّذ `brew install libheif` ثم أعد الفحص باستخدام `uv run --frozen docwen doctor --json`. تعتمد الإتاحة على البناء والمحرّك الفعليين.

## العمل من الواجهة الرسومية

1. أضف ملفًا، أو اختر مجموعة ملفات للعمل عليها دفعة واحدة.
2. اختر صيغة الإخراج المناسبة للملف.
3. اضبط خيارات الصور أو التعرف على النص أو القالب عند الحاجة.
4. اضغط زر العملية وراجع مكان حفظ النتيجة.

يمكنك مشاهدة اختيار القالب وخيارات إنشاء المستند هنا:

![إنشاء مستند من Markdown](../assets/screenshots/markdown-generation-en.png)

خيارات محاذاة ملاحظات المرفقات في تنسيق المستندات الرسمية (تظهر الصورة واجهة إنجليزية):

![إعداد محاذاة معلومات المرفقات في المستندات الرسمية](../assets/screenshots/official-documents-en.png)


النتيجة ليست دائمًا مطابقة للأصل. تأكد من تنسيق الجداول والمعادلات والنص المعترف به عبر OCR قبل استخدام الملف النهائي.

عند بدء المهمة تظهر خيارات الطلب الفعلي. ابحث عن القوالب بالاسم أو بوصف الاستخدام؛ تعرض قائمة «المستخدمة مؤخرًا في هذه الجلسة» قوالب المهام التي بدأت، دون تغيير القالب الافتراضي. لدمج الجداول راجع [النطاق والأوضاع ومراجعة النتائج](table-merge.md).

## استخدام سطر الأوامر

تستخدم أمثلة CLI اسم الأمر `docwen`. من مجلد المصدر أضف قبله `uv run --frozen`؛ وفي الحزمة استخدم `DocWenCLI.exe` على Windows أو `./DocWenCLI` على Linux/macOS.

تتيح واجهة CLI فحص الملف واختيار العملية قبل الكتابة. هذه الأوامر مخصصة لمن يستخدم الطرفية أو برامج الأتمتة:

```bash
# افحص تنسيق الملف والعمليات الممكنة
docwen inspect <file> [--json]
# قائمة التنسيقات والقوالب والتحسينات المتاحة
docwen resources list formats --json
docwen resources list optimizations --json
# اعرض المعلمات المتاحة للتحويل
docwen schema convert
# جرّب تحديد المسار دون كتابة ملفات
docwen convert sample.docx --to md --output-dir output --dry-run --json
# ابدأ التحويل الفعلي
docwen convert sample.docx --to md --output-dir output
```

يمكن لأنظمة التكامل استخدام `serve --stdio` عبر **Machine Protocol v2** بدلًا من قراءة نص المخرجات المصمم للمستخدم.

| الأمر | الوظيفة |
| --- | --- |
| `gui open [ABSOLUTE_FILE]` | فتح واجهة DocWen وإضافة ملف موجود اختياريًا |
| `inspect <file> [--json]` | معرفة التنسيق والعمليات المتاحة |
| `resources list formats --json` | استعراض تنسيقات التحويل |
| `resources list optimizations --json` | استعراض التحسينات |
| `schema convert` | عرض خيارات التحويل |
| `serve --stdio` | بدء الاتصال المحلي الخاص بالتكامل |

للتفاصيل راجع [دليل سطر الأوامر](../cli.md) و[عقد البروتوكول](../specs/machine-protocol-v2.md).

من أمثلة العمل المباشر في سطر الأوامر: لا يُنشأ تقرير التدقيق إلا عند طلب `--report` صراحةً، ويجب فحص موارد التحسين المتاحة قبل اختيارها.

```bash
docwen resources list optimizations --json
docwen convert report.docx --to md --output-dir exports --optimization gongwen
docwen validate input.md --check typo --report report.json
```

## صيغة Markdown والامتدادات

يمكن استخدام DocWen مع ملفات Markdown العادية دون إضافات. أما صيغة Structural Tables والجداول المدمجة وتسميات Number Suite والعناوين H7–H9 والحواشي الختامية المميزة فهي **امتدادات اختيارية معطّلة افتراضيًا**.

عند التحويل من Markdown إلى Word، فعّل امتدادات **الإدخال** التي تستخدمها ملاحظاتك من إعدادات Markdown. وعند التحويل العكسي، فعّل امتدادات **الإخراج** المطلوبة. قد تتغير المسافات أو كتابة العلامات أثناء التحويل، لذا احتفظ بالأصل وراجع النتيجة.

## Obsidian والإضافات المرافقة

يدعم ملحق DocWen Assistant سطح المكتب على Windows وLinux، لكنه لا يعمل على نسخة Obsidian للأجهزة المحمولة. كما لا يدعم DocWen Assistant نظام macOS حاليًا.

يمكنك تثبيت [DocWen Assistant](https://github.com/ZHYX91/obsidian-docwen-assistant) لتحويل الملفات أو تدقيقها من داخل Obsidian. اختر إصدارًا من [صفحة الإصدارات المنشورة](https://github.com/ZHYX91/obsidian-docwen-assistant/releases) وتحقق من توافقه مع DocWen المثبت على جهازك، بدلًا من الاعتماد على رقم إصدار ثابت هنا. ويمكن استخدام [docwen-openclaw](https://github.com/ZHYX91/docwen-openclaw) لتشغيل أدوات DocWen من OpenClaw؛ تحقق من [الإصدارات المنشورة له](https://github.com/ZHYX91/docwen-openclaw/releases) أيضًا. **كلتا الإضافتين تحتاجان إلى تثبيت DocWen متوافق على الجهاز**؛ الإضافة وحدها لا تثبّت البرنامج.

## الأمان والخصوصية

يعالج DocWen الملفات على جهازك، ولا يوفر رفع الملفات أو القياس عن بُعد أو التنزيل عبر الإنترنت أو خدمة شبكية. يقيّد الوصول إلى الشبكة لمكوّنات Python الخاصة به. إذا تعذّر تفعيل هذه الحماية، فلن يواصل DocWen التشغيل.

البرامج التي تُشغّل بصورة منفصلة، مثل Word/Office وWPS وLibreOffice، بما فيها العمليات المساعدة، تتبع إعدادات الشبكة الخاصة بها وبنظام التشغيل. هذا القيد لا يشملها، لذا راجع إعداداتها أيضًا عند التعامل مع ملفات حساسة.

تشرح [وثائق البنية](../architecture.md) كيفية عمل الحماية وحدودها. وإذا تعذّر تشغيل البرنامج أو فشل التحويل، فراجع [دليل حل المشكلات](../maintenance/troubleshooting.md).

## الأسئلة والمساعدة

تدعم CLI `validate` ملفات `.md` و`.docx` مباشرةً. ويمكن تدقيق صيغ Word القديمة (`.doc` و`.wps` و`.rtf` و`.odt`) بعد تحويلها إلى DOCX ببرنامج Office محلي متوافق؛ تحقق من توفر برنامج التحويل أولًا. العملية للقراءة فقط ما لم يُطلب إخراج تقرير صراحةً.

إن فشل التحويل، افحص نوع الملف وحالة برامج Word أو WPS أو LibreOffice المستخدمة. جرّب مجلد إخراج جديدًا؛ ولا تُعد تشغيل عملية كتابة غير مؤكدة قبل مراجعة النتائج الموجودة.

لطرح سؤال أو تقديم بلاغ، استخدم [مشكلات DocWen](https://github.com/ZHYX91/docwen/issues). راجع [سجل التغييرات](../CHANGELOG.md) لمعرفة ما تغير في كل إصدار.

## الترخيص

DocWen برنامج مفتوح المصدر بترخيص AGPL-3.0. راجع [ملف الترخيص](../../LICENSE).
