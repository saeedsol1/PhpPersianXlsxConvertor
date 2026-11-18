<div dir="rtl" align="right">

# 🔄 راهنمای نصب مبدل فایل (CSV / TXT / Excel)

این برنامه یک صفحهٔ وب ساده است که فایل‌های **CSV، TXT، XLSX، XLS و ODS** را به **XLSX، XLS، ODS، CSV، HTML و PDF** تبدیل می‌کند. فایل‌های تبدیل‌شده یک ساعت نگهداری می‌شوند و بعد خودکار حذف می‌شوند.

---

## 📋 پیش‌نیازها

| ابزار | کاربرد | لینک دانلود |
|:--|:--|:--|
| **XAMPP** | وب‌سرور آپاچی و PHP | [xampp-windows-x64-8.5.12](https://sourceforge.net/projects/xampp/files/XAMPP%20Windows/8.5.12/xampp-windows-x64-8.5.12-0-VS17-installer.exe) |
| **Composer** | نصب کتابخانهٔ PhpSpreadsheet | [Composer-Setup.exe](https://getcomposer.org/Composer-Setup.exe) |

> 💡 برای این مبدل فقط سرویس **Apache** لازم است؛ بقیهٔ سرویس‌های XAMPP را فعال نکنید.

---

## 1️⃣ نصب XAMPP

1. فایل نصب XAMPP را از لینک بالا دانلود و نصب کنید (مثلاً در مسیر `D:\xampp`).
2. بعد از نصب، **XAMPP Control Panel** را باز کنید. هنوز چیزی را Start نکنید؛ ابتدا مراحل بعد را انجام دهید.

---

## 2️⃣ نصب Composer

1. فایل `Composer-Setup.exe` را دانلود و اجرا کنید.
2. در مرحله‌ای که مسیر PHP را می‌پرسد، **مسیر فایل `php.exe`** خودتان را بدهید. مثلاً:

<div dir="ltr" align="left">

```text
D:\xampp\php\php.exe
```

</div>

3. نصب را تمام کنید. بعد از نصب، پنجرهٔ cmd را اگر باز بود ببندید و دوباره باز کنید.

---

## 3️⃣ ساخت پوشهٔ برنامه

1. داخل پوشهٔ `htdocs` یک پوشهٔ دلخواه بسازید، مثلاً `excel`:

<div dir="ltr" align="left">

```text
D:\xampp\htdocs\excel
```

</div>

2. فایل‌های برنامه را از GitHub بگیرید: در صفحهٔ مخزن روی دکمهٔ سبز **Code** و سپس **Download ZIP** بزنید.
3. فایل زیپ را **Extract** کنید و محتویات آن را داخل پوشهٔ `excel` قرار دهید.

> ⚠️ زیپ GitHub همهٔ فایل‌ها را داخل یک پوشهٔ اضافه (مثلاً `نام-مخزن-main`) می‌گذارد. فایل‌ها را از داخل آن پوشه بیرون بیاورید تا `index.php` **مستقیماً** داخل پوشهٔ `excel` دیده شود، نه داخل یک پوشهٔ دیگرِ تودرتو:
> `D:\xampp\htdocs\excel\index.php` ✅

---

## 4️⃣ نصب کتابخانه‌ها

1. پوشهٔ `D:\xampp\htdocs\excel` را در File Explorer باز کنید.
2. روی **نوار آدرس** کلیک کنید، عبارت `cmd` را تایپ کنید و **Enter** بزنید تا پنجرهٔ cmd در همین پوشه باز شود.
3. دستور زیر را کپی و اجرا کنید:

<div dir="ltr" align="left">

```bash
composer require phpoffice/phpspreadsheet
```

</div>

4. منتظر بمانید دانلود تمام شود. پوشهٔ `vendor` و فایل‌های `composer.json` و `composer.lock` داخل پوشه ساخته می‌شوند.

### 📄 فقط اگر خروجی PDF می‌خواهید

دستور زیر را هم اجرا کنید:

<div dir="ltr" align="left">

```bash
composer require mpdf/mpdf
```

</div>

> 💡 اگر Composer دربارهٔ نبودن یک افزونه (مثل `gd` یا `zip`) خطا داد، فایل `D:\xampp\php\php.ini` را باز کنید، `;` اول خط `extension=gd` یا `extension=zip` را بردارید، ذخیره کنید و دستور را دوباره اجرا کنید.

---

## 5️⃣ اجرا

1. در XAMPP Control Panel کنار **Apache** روی **Start** بزنید.
2. این آدرس را در مرورگر باز کنید:

<div dir="ltr" align="left">

```text
http://localhost/excel/
```

</div>

صفحهٔ اصلی مبدل نمایش داده می‌شود. 🎉

---

## 🛠️ عیب‌یابی سریع

| مشکل | راه‌حل |
|:--|:--|
| Apache روشن نمی‌شود | احتمالاً پورت 80 را برنامهٔ دیگری (مثل IIS یا Skype) گرفته است. آن برنامه را ببندید یا پورت Apache را عوض کنید. |
| خطای `vendor/autoload.php` | دستور `composer require` در پوشهٔ درست (کنار `index.php`) اجرا نشده است. |
| `composer` شناخته نمی‌شود | cmd را ببندید و دوباره باز کنید؛ اگر نشد، Composer را دوباره نصب کنید. |
| گزینهٔ PDF غیرفعال است | دستور `composer require mpdf/mpdf` اجرا نشده است. |

</div>
