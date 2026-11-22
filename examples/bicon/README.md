# Bicon — مثال التحقق

مشروع تدريب. الـ kit اتبنى عكسياً من مخرجاته المعتمدة، والمثال ده بيتأكد إن السكريبتات بتطلّع نفس الأرقام.

```bash
cd ../..
python3 schedule.py examples/bicon/project.json /tmp/bicon
python3 update.py   examples/bicon/project.json /tmp/bicon
python3 cost.py     examples/bicon/project.json /tmp/bicon
python3 viewer.py   examples/bicon/project.json /tmp/bicon
python3 verify.py   examples/bicon/project.json /tmp/bicon
```

النتيجة المتوقعة:

| البند | القيمة |
|---|---|
| الأنشطة | 106، كلها مطابقة للمرجع (التواريخ، الـ float، المنطق، الأطقم) |
| نهاية البيز لاين / المتوقعة | 15-Apr-2027 / 13-May-2027، والتأخير 28 يوم تقويمي |
| العناصر المربوطة | 548 من 548 |
| BAC | 11,279,094 EGP |
| PV / EV / AC بتاريخ 30-Nov-2026 | 9.35 / 8.82 / 13.04 مليون (الـ AC هنا ASSUMED) |
| CPI / SPI / EAC | 0.676 / 0.943 / 16.68 مليون |

`test_cost_sheet.xlsx` اتعمل من شيت Unit Rates في المرجع، عشان نجرّب عليه `boq_rates.py`. فيه أعمدة سعر بيع مقصودة، والأداة لازم تنبّه عليها.

اختلافات مقصودة عن المرجع:
- العمالة الفعلية في المرجع كانت نسبة تخيلية من المخطط. الـ kit بياخدها من الأطقم المكتوبة في شيت الـ Update.
- الـ Cost Report في المرجع كان بيحسب أيام العمل في PV% / EV% بـ weekend code 11، وده بيعتبر الأحد إجازة. الـ kit بيستخدم إجازة التقويم الفعلية، فالمعادلات في Excel بقت مطابقة لحسابات Python.
