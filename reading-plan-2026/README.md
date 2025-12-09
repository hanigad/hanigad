# قراءة الكتاب المقدس كاملا في عام واحد (بالموضوعات) — 2026

ملفّات هذه الخطة تحتوي مراجع القراءة اليومية (365 يومًا) مقسومة إلى 4 أجزاء يومية: العهد القديم، العهد الجديد، مزمور، أمثال.

الملفّات المتوفرة حالياً في هذا PR:
- reading-plan-2026/reading_plan_2026_sample_7days_corrected.csv — CSV UTF-8 عيّنة 7 أيام مصحّحة.
- reading-plan-2026/reading_plan_2026_sample_7days_corrected.json — JSON مهيأ للبوتات والتوزيع (حقل message مضمّن).
- reading-plan-2026/export_to_json.bas — كود VBA لتصدير الورقة Plan إلى JSON داخل ملف .xlsm.
- README.md — هذا الملف.

ملاحظات:
- لا يتم تضمين ملف .xlsm الثنائي في الشجرة؛ سيتم رفعه كـ Release asset لاحقًا.
- روابط النصوص تشير إلى ترجمة فاندايك العربية على ebible.org.

كيفية الاستخدام السريع:
1. افتح Excel وأنشئ ورقة Settings: ضع B1 = 2026.
2. افتح ورقة Plan وألصق CSV (A..L).
3. ألصق كود export_to_json.bas في Module ثم احفظ كـ .xlsm.
4. شغّل الماكرو ExportPlanToJSON لتوليد JSON محليًا.

حقوق:
النصوص مأخوذة من ترجمة Van Dyck على ebible.org — المرجع: https://ebible.org/arb-vd/
