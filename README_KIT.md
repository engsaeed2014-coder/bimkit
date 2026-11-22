# BIM kit — 3D / 4D / 5D

السكريبتات اللي بتشتغل بيها الـ skills الثلاثة:
`drawings-to-3d-model` (3D) ← `bim-4d-schedule` (4D) ← `bim-5d-cost` (5D).
الـ 3D بيطلّع `data.json` (عناصر الموديل). الـ kit بياخده ويبني عليه البرنامج والتكلفة وصفحة العرض.

المتطلبات: Python 3.10+، `openpyxl`، `shapely` (للـ viewer)، وLibreOffice (لـ verify بس).

```
pip install openpyxl shapely --break-system-packages
```

## الملفات

| الملف | الشغل |
|---|---|
| `project.json` | تعريف المشروع: التقويم، الأدوار، الزونات، الخطوات والأطقم، المنطق، قواعد التكلفة. **ده الملف الوحيد اللي بيتغيّر من مشروع لمشروع.** |
| `schedule.py` | البرنامج الأساسي (Baseline): توليد الأنشطة والمنطق من project.json أو استيراد XER، حساب CPM، ربط كل عنصر بأنشطته. |
| `update.py` | يقرا شيت التحديث (Update.xlsx) ويطلّع `feed.json` للـ 4D: الفعلي، المتوقع، التأخير، العمالة اليومية. |
| `boq_rates.py` | يقرا شيت التكاليف بأي شكل ويطلّع `rates.json` (سعر التكلفة = التكلفة ÷ الكمية). |
| `cost.py` | الـ 5D: ميزانية كل عنصر وكل نشاط، PV / EV / AC، CPI / SPI، EAC، وملفين Excel. |
| `viewer.py` | يبني صفحة واحدة `out/index.html` فيها تابات 3D و4D و5D من القوالب المعتمدة. |
| `verify.py` | فحص قبل التسليم (المنطق، الربط، الميزانية، Excel بدون أخطاء، مفيش أسعار بيع). |
| `xer.py`, `progxlsx.py`, `common.py` | قراءة وكتابة XER، ملفات Excel الخاصة بالبرنامج، والتقويم وحساب الـ CPM. |
| `templates/` | صفحة العرض المعتمدة (shell + 3D + 4D/5D). متتعدلش غير لو هنغيّر شكل الصفحة نفسها. |
| `adapters/reveal4d.py` | محوّل REVEAL: موديل reveal-4d + الماستر المعتمد → `data.json` و`activities.json` و`links.json`. |
| `examples/bicon/` | مثال كامل شغّال (مشروع تدريب) للمقارنة. |
| `examples/reveal-eb/` | `project.json` لبلوك EB. ملفات البيانات بتتعمل بالمحوّل ومش بتترفع، لأنها من ملفات المشروع السرية. |

## التشغيل

```bash
P=project.json; O=out

# 4D
python3 schedule.py $P $O                 # programme.json + Programme.xlsx + Baseline XER + Update.xlsx (أول مرة)
#   أو من برنامج موجود:  python3 schedule.py $P $O --xer Baseline.xer
python3 update.py   $P $O [Update.xlsx]   # feed.json + تحديث Programme.xlsx

# 5D
python3 boq_rates.py inspect CostSheet.xlsx              # الشيتات
python3 boq_rates.py inspect CostSheet.xlsx "Sheet"      # الأعمدة وتخمين DESC / QTY / COST
python3 boq_rates.py extract CostSheet.xlsx "Sheet" 4 120 B D F rates.json C
python3 cost.py $P $O [--ac Cost_Report.xlsx]            # Budget_Baseline.xlsx + Cost_Report.xlsx + cost.json

# الصفحة والفحص
python3 viewer.py $P $O                   # out/index.html
python3 verify.py $P $O
```

كل أمر بيطبع ملخص قصير: الـ BAC، العناصر اللي ملهاش قاعدة سعر (لازم تبقى 0)، والأسعار المستخدمة.

## برنامج معتمد جاهز (خرسانة + تشطيبات)

لو المشروع عنده ماستر معتمد بتواريخه وبالـ cost loading بتاعه، الـ kit بياخده زي ما هو من غير ما يولّد برنامج:

```bash
python3 adapters/reveal4d.py --reveal4d ../reveal-4d --master master.xlsx --area EB --out examples/reveal-eb
python3 schedule.py examples/reveal-eb/project.json out      # inputs.activities + inputs.links
python3 update.py   examples/reveal-eb/project.json out --data-date 2026-11-22 --assume-baseline-progress
python3 cost.py     examples/reveal-eb/project.json out      # budget mode: activity
python3 viewer.py   examples/reveal-eb/project.json out
```

- `activities.json`: فيه لكل نشاط الكود، والاسم، والدور، والزون، والبداية والنهاية، وكمان `budget` = الكمية، والوحدة، والسعر، والتكلفة، والمصدر.
- التواريخ بتتاخد زي ما هي من غير CPM. عشان كده الـ float مش بيتحسب.
- `links.json`: كل عنصر في الموديل وأنشطته. النشاط اللي مالوش عنصر بيدخل الـ 4D والـ 5D على مستوى النشاط بس.
- العنصر ممكن يكون مجمّع من قطع كتير (`geom.parts`). الصفحة بتدمجهم في شكل واحد، وده بيخلي بلوك فيه 6 عمارات يتعرض بسرعة.
- `--assume-baseline-progress`: أي نشاط ملوش سجل في شيت الـ Update بيمشي على تواريخ البيز لاين، والتقدم بيبقى ASSUMED لحد ما سجلات الموقع توصل.
- لما الميزانية تكون جاية من الأنشطة، `cost.budget_note` بيوضّح مصدرها في كل الملفات.
- لو فيه ماستر متحدّث من سجلات الموقع (IR / CPR): `adapters/reveal_progress.py` بيملا شيت الـ Update منه بتاريخ البيانات بتاعه. النشاط اللي خلص قبل تاريخ البيانات بيبقى 100%، واللي شغال بتتحسب نسبته من الأيام اللي فاتت منه (ASSUMED)، والباقي بيبقى forecast. وبعدها `update.py` بيقارن الشيت ده بالبيز لاين المعتمد.

## project.json — أهم الأجزاء

- `calendar.off_weekdays`: أيام الإجازة بأرقام Python (الاثنين = 0 … الجمعة = 4، الأحد = 6). `holidays`: إجازات رسمية `YYYY-MM-DD`.
- `schedule.start` و`contract_end`. المدد: `base` لكل خطوة × `duration_scale` (معامل K)، والتقريب `ceil`. أو `dur` مباشرة للخطوة.
- `schedule.zones`: تقسيم العناصر على الزونات بالإحداثيات (`x_min` / `x_max`، والحد الأعلى مش داخل في الزون).
- `schedule.floors` و`floor_parts`: الأجزاء في كل دور (`SLB` بلاطة، `VERT` أعمدة وحوائط). الدور اللي ملوش بلاطة، زي البدروم فوق اللبشة، يتكتب `["VERT"]`.
- `schedule.sequence` و`weights`: ترتيب الخطوات جوه كل مجموعة، ووزن كل خطوة في التقدم والميزانية. القيم الافتراضية: لبشة .2 / .4 / .4، بلاطة .35 / .35 / .30، أعمدة وحوائط .75 / .25.
- `schedule.logic`: lag الزون التالية (SS)، والفواصل بين صب اللبشة والأعمدة، والبلاطة والأعمدة.
- `schedule.steps`: اسم كل خطوة (عربي وإنجليزي)، المدة، والطاقم لكل يوم.
- `schedule.xer_map`: عند استيراد XER، ربط كود النشاط بـ floor / zone / step. لو أسماء الأنشطة على نمط الـ kit (`GF - Wing A - Slab formwork`) الربط بيتعمل لوحده.
- `cost.rules`: أول قاعدة تنطبق على العنصر هي اللي بتتطبق. `where` بياخد type / mark / group / floor. `rate` بياخد أول كلمات وصف البند من شيت التكلفة، أو قائمة بنود تتجمع، أو `{by:"t_m", table:{سمك: وصف}}` للبلاطات حسب السمك. `basis` إما `volume_m3` (الافتراضي) أو `area_m2`. `also` بنود إضافية زي العادية والعزل بالمتر المسطح.
- `cost.ac_factors`: AC تخيلي لكل خطوة (ASSUMED). يتشال أول ما التكلفة الفعلية تيجي من الموقع.
- `viewer.offset`: إزاحة الموديل في صفحة الـ 4D. `rings_floor`: الدور اللي بيتاخد منه حدود المبنى.

## قواعد ثابتة

- الـ 5D موديل **تكلفة**. سعر الميزانية = سعر شيت التكلفة. `EV = سعر الميزانية × الكمية المنفذة`. أسعار البيع أو العميل أو قيمة المقايسة ممنوعة في أي مخرج، و`verify.py` بيدور عليها.
- كل رقم عليه tag: READ / CALC / ASSUMED. أي حاجة ناقصة تتكتب كفجوة واضحة، مفيش تخمين.
- النشاط اللي لسه مبدأش ومواعيده في الشيت أبكر من اللي السوابق تسمح بيه، بيتأخر لوحده عشان يحترم المنطق ومدته بتفضل زي ما هي. عشان تلغي ده: `--keep-sheet-forecast`.
- التأخير في `feed.json` بالأيام التقويمية = آخر نهاية متوقعة − آخر نهاية في البيز لاين (زي Finish slip في شيت الـ Update).
- العمالة الفعلية = الأطقم المكتوبة في شيت الـ Update لحد تاريخ البيانات. ممكن تضيف `update.actual_labour_factor` (رقم، أو قيمة لكل خطوة) كـ ASSUMED لحد ما أرقام الموقع توصل.

## حدود الـ kit

- الخرسانة بس (بنود بالـ m³ أو الـ m²). الحديد بالطن، والشدة بالـ m²، والكهروميكانيكا، والتشطيبات، والمصاريف غير المباشرة، محتاجين قواعد وبنود إضافية.
- سعر ميزانية واحد لكل قاعدة، من غير زيادة أسعار مع الوقت.
- صفحة العرض بتحمّل three.js وSheetJS من CDN، يعني محتاجة نت.
