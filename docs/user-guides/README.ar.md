# DocWen

<p align="center">
  <img src="https://raw.githubusercontent.com/ZHYX91/docwen/main/assets/icon.svg" alt="DocWen logo" width="120">
</p>

[English](https://github.com/ZHYX91/docwen/blob/main/README.md) · [简体中文](https://github.com/ZHYX91/docwen/blob/main/docs/user-guides/README.zh-CN.md) · [繁體中文](https://github.com/ZHYX91/docwen/blob/main/docs/user-guides/README.zh-TW.md) · [Deutsch](https://github.com/ZHYX91/docwen/blob/main/docs/user-guides/README.de-DE.md) · [Français](https://github.com/ZHYX91/docwen/blob/main/docs/user-guides/README.fr-FR.md) · [Español](https://github.com/ZHYX91/docwen/blob/main/docs/user-guides/README.es-ES.md) · [Português](https://github.com/ZHYX91/docwen/blob/main/docs/user-guides/README.pt-BR.md) · [Русский](https://github.com/ZHYX91/docwen/blob/main/docs/user-guides/README.ru-RU.md) · [日本語](https://github.com/ZHYX91/docwen/blob/main/docs/user-guides/README.ja-JP.md) · [한국어](https://github.com/ZHYX91/docwen/blob/main/docs/user-guides/README.ko-KR.md) · [Tiếng Việt](https://github.com/ZHYX91/docwen/blob/main/docs/user-guides/README.vi-VN.md) · [العربية](https://github.com/ZHYX91/docwen/blob/main/docs/user-guides/README.ar.md)

يساعدك DocWen على تحويل المستندات والجداول والصور **على جهازك**، دون رفع الملفات إلى خدمة تحويل على الإنترنت. يمكنك استخدام الواجهة الرسومية للأعمال اليومية أو سطر الأوامر لتنفيذ عمليات متكررة.

## ما الذي يمكن فعله؟

- تحويل ملفات Word إلى Markdown والعكس، مع الحفاظ على العناصر التي تدعمها طريقة التحويل المختارة.
- تصدير جداول Excel إلى Markdown وإنشاء جداول من ملفات Markdown.
- استخراج النص من PDF وXPS وOFD؛ والتعرف على النص داخل الصور محليًا عند توفر OCR.
- دمج ملفات PDF أو تقسيمها، وتحويل الصور إلى تنسيقات أخرى.
- اختيار قوالب Word أو Excel الخاصة بك، وفحص النصوص بقواميس وقواعد مخصّصة.

لا يدعم كل تنسيق جميع عمليات التحويل. يمكنك فحص الملف والعمليات المتاحة أولًا؛ راجع [جدول القدرات](../capabilities.md).

## البدء بسرعة

يمكن تنزيل الإصدارات الجاهزة من [صفحة إصدارات DocWen](https://github.com/ZHYX91/docwen/releases). الحزم الرسمية المعلنة هي Windows x64 وUbuntu 24.04 x64؛ أما الأنظمة الأخرى فتحتاج إلى تشغيل من المصدر ولا تُعد ضمن دعم الحزم الرسمية.

على Windows، افتح `DocWen.exe`. في حزمة Linux الرسومية، شغّل `./DocWen` بعد استخراج الأرشيف بالكامل. تتوفر حزمة CLI منفصلة.

للتشغيل من المصدر، تحتاج إلى Python 3.12 وأداة uv 0.12.0:

```bash
git clone https://github.com/ZHYX91/docwen.git
cd docwen
uv sync --frozen --all-extras
docwen-gui
```

ليست حزمة macOS جاهزة رسميًا، ولا تدعم عمليات التحويل الأساسية على macOS حاليًا. لا تضمن حزمة Ubuntu العمل على توزيعات أو معالجات Linux الأخرى.

## العمل من الواجهة الرسومية

1. أضف ملفًا، أو اختر مجموعة ملفات للعمل عليها دفعة واحدة.
2. اختر صيغة الإخراج المناسبة للملف.
3. اضبط خيارات الصور أو التعرف على النص أو القالب عند الحاجة.
4. اضغط زر العملية وراجع مكان حفظ النتيجة.

يمكنك مشاهدة اختيار القالب وخيارات إنشاء المستند هنا:

![إنشاء مستند من Markdown](../assets/screenshots/markdown-generation-en.png)

وتستطيع إضافة القوالب وترتيبها في الإعدادات:

![إدارة القوالب](../assets/screenshots/template-management-en.png)

النتيجة ليست دائمًا مطابقة للأصل. تأكد من تنسيق الجداول والمعادلات والنص المعترف به عبر OCR قبل استخدام الملف النهائي.

## استخدام سطر الأوامر

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

## Markdown وObsidian

يمكن استخدام DocWen مع ملفات Markdown العادية دون إضافات. أما صيغة Structural Tables والجداول المدمجة وتسميات Number Suite والعناوين H7–H9 والحواشي الختامية المميزة فهي **امتدادات اختيارية معطّلة افتراضيًا**.

عند التحويل من Markdown إلى Word، فعّل امتدادات **الإدخال** التي تستخدمها ملاحظاتك من إعدادات Markdown. وعند التحويل العكسي، فعّل امتدادات **الإخراج** المطلوبة. قد تتغير المسافات أو كتابة العلامات أثناء التحويل، لذا احتفظ بالأصل وراجع النتيجة.

## Obsidian والإضافات المرافقة

يدعم ملحق DocWen Assistant سطح المكتب على Windows وLinux، لكنه لا يعمل على نسخة Obsidian للأجهزة المحمولة. كما لا يدعم DocWen Assistant نظام macOS حاليًا.

يمكنك تثبيت [DocWen Assistant](https://github.com/ZHYX91/obsidian-docwen-assistant) لتحويل الملفات أو تدقيقها من داخل Obsidian. ويمكن استخدام [docwen-openclaw](https://github.com/ZHYX91/docwen-openclaw) لتشغيل أدوات DocWen من OpenClaw. **كلتا الإضافتين تحتاجان إلى تثبيت DocWen متوافق على الجهاز**؛ الإضافة وحدها لا تثبّت البرنامج.

## الأمان والخصوصية

يعالج DocWen ملفاتك محليًا. عند تشغيله، يمنع حارس الشبكة داخل عملية Python التابعة له **جميع عمليات DNS وحل الأسماء**، إضافة إلى اتصالات `AF_INET/AF_INET6` بما فيها `bind` و`connect` و`connect_ex` و`sendto` و`sendmsg`.

هذا الحظر لا يغطي تلقائيًا العمليات الخارجية. يمكن تشغيل Word أو WPS أو LibreOffice عبر عملية Office مساعدة مخصّصة، وتظل هذه البرامج خاضعة لسياسات الشبكة الخاصة بها. إذا احتجت إلى منع وصولها للشبكة تمامًا، فاستخدم عزل على مستوى نظام التشغيل. هذا ينطبق على أنظمة Windows وUnix. لا تعتبر عبارة «محلي» ضمانًا بأن كل برامج Office الخارجية محظورة الشبكة.

قبل نشر سجلات الأخطاء، احذف نصوص المستندات ومسارات الملفات ومعلومات الدخول. راجع [الوثائق التقنية](../architecture.md) و[متطلبات التشغيل](../external-dependencies.md).

## الأسئلة والمساعدة

إن فشل التحويل، افحص نوع الملف وحالة برامج Word أو WPS أو LibreOffice المستخدمة. جرّب مجلد إخراج جديدًا؛ ولا تُعد تشغيل عملية كتابة غير مؤكدة قبل مراجعة النتائج الموجودة.

لطرح سؤال أو تقديم بلاغ، استخدم [مشكلات DocWen](https://github.com/ZHYX91/docwen/issues). راجع [سجل التغييرات](../CHANGELOG.md) لمعرفة ما تغير في كل إصدار.

## الترخيص

DocWen برنامج مفتوح المصدر بترخيص AGPL-3.0. راجع [ملف الترخيص](../../LICENSE).
