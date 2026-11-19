<div dir="rtl" align="right">
# 🔄 مبدل فایل (CSV / TXT / Excel)

این برنامه یک صفحهٔ وب ساده است که فایل‌های **CSV، TXT، XLSX، XLS و ODS** را به **XLSX، XLS، ODS، CSV، HTML و PDF** تبدیل می‌کند. فایل‌های تبدیل‌شده یک ساعت نگهداری می‌شوند و بعد خودکار حذف می‌شوند.

---

## 📋 پیش‌نیازها

| ابزار | کاربرد | لینک دانلود |
|:--|:--|:--|
| **XAMPP** | وب‌سرور آپاچی و PHP | [صفحهٔ دانلود XAMPP](https://www.apachefriends.org/download.html) |
| **Composer** | نصب کتابخانه‌های پروژه | [Composer-Setup.exe](https://getcomposer.org/Composer-Setup.exe) |

> 💡 برای این مبدل فقط سرویس **Apache** لازم است؛ بقیهٔ سرویس‌های XAMPP را فعال نکنید.

> 💡 مطمئن شوید نسخهٔ XAMPPی که دانلود می‌کنید از PHP 8.1 یا بالاتر استفاده می‌کند (نسخه‌های جدید PHP در صفحهٔ رسمی XAMPP موجود است).

---

## 1️⃣ نصب XAMPP

1. فایل نصب XAMPP را از لینک بالا دانلود و نصب کنید (مثلاً در مسیر `D:\xampp`).
2. بعد از نصب، **XAMPP Control Panel** را باز کنید. هنوز چیزی را Start نکنید؛ ابتدا مراحل بعد را انجام دهید.

---

## 2️⃣ نصب Composer

1. فایل `Composer-Setup.exe` را دانلود و اجرا کنید.
2. در مرحله‌ای که مسیر PHP را می‌پرسد، **مسیر فایل `php.exe`** خودتان را بدهید. مثلاً:


D:\xampp\php\php.exe
نصب را تمام کنید. بعد از نصب، پنجرهٔ cmd را اگر باز بود ببندید و دوباره باز کنید.

3️⃣ ساخت پوشهٔ برنامه
داخل پوشهٔ htdocs یک پوشهٔ دلخواه بسازید، مثلاً excel:


D:\xampp\htdocs\excel
فایل‌های برنامه را از GitHub بگیرید: در صفحهٔ مخزن روی دکمهٔ سبز Code و سپس Download ZIP بزنید.

فایل زیپ را Extract کنید و محتویات آن را داخل پوشهٔ excel قرار دهید.

⚠️ زیپ GitHub همهٔ فایل‌ها را داخل یک پوشهٔ اضافه (مثلاً نام-مخزن-main) می‌گذارد. فایل‌ها را از داخل آن پوشه بیرون بیاورید تا index.php مستقیماً داخل پوشهٔ excel دیده شود، نه داخل یک پوشهٔ دیگرِ تودرتو:
D:\xampp\htdocs\excel\index.php ✅

4️⃣ نصب کتابخانه‌ها
پوشهٔ D:\xampp\htdocs\excel را در File Explorer باز کنید.

روی نوار آدرس کلیک کنید، عبارت cmd را تایپ کنید و Enter بزنید تا پنجرهٔ cmd در همین پوشه باز شود.

دستور زیر را کپی و اجرا کنید:

bash
composer install
این دستور هر دو کتابخانهٔ موردنیاز پروژه را بر اساس composer.json نصب می‌کند:

phpoffice/phpspreadsheet — خواندن و نوشتن فایل‌های Excel

mpdf/mpdf — تولید خروجی PDF

منتظر بمانید دانلود تمام شود. پوشهٔ vendor و فایل‌های composer.json و composer.lock داخل پوشه ساخته می‌شوند.

💡 اگر Composer دربارهٔ نبودن یک افزونه (مثل gd یا zip) خطا داد، فایل D:\xampp\php\php.ini را باز کنید، ; اول خط extension=gd یا extension=zip را بردارید، ذخیره کنید و دستور را دوباره اجرا کنید.

💡 اگر روی کامپیوتر شما از قبل vendor وجود دارد و فقط می‌خواهید نسخه‌ها را به‌روز کنید، از composer update استفاده کنید.

5️⃣ اجرا
در XAMPP Control Panel کنار Apache روی Start بزنید.

این آدرس را در مرورگر باز کنید:


http://localhost/excel/
صفحهٔ اصلی مبدل نمایش داده می‌شود. 🎉

🛠️ عیب‌یابی سریع
مشکل	راه‌حل
Apache روشن نمی‌شود	احتمالاً پورت 80 را برنامهٔ دیگری (مثل IIS یا Skype) گرفته است. آن برنامه را ببندید یا پورت Apache را عوض کنید.
خطای vendor/autoload.php	دستور composer install در پوشهٔ درست (کنار index.php) اجرا نشده است.
composer شناخته نمی‌شود	cmd را ببندید و دوباره باز کنید؛ اگر نشد، Composer را دوباره نصب کنید.
گزینهٔ PDF غیرفعال است	مطمئن شوید mpdf/mpdf در vendor نصب شده است (composer install را دوباره اجرا کنید).
خطای Class "Mpdf\Mpdf" not found	اجرای composer install ناقص مانده است یا vendor/autoload.php در فایل PHP بارگذاری نشده است.
📁 ساختار پروژه
text
excel/
├── index.php
├── composer.json
├── composer.lock
├── vendor/
├── LICENSE
└── README.md
پوشهٔ vendor و فایل composer.lock توسط Composer ساخته می‌شوند و نباید دستی ویرایش شوند.

📄 لایسنس
این پروژه تحت مجوز GNU General Public License v2.0 only منتشر شده است. فایل LICENSE را ببینید.

text
This program is free software; you can redistribute it and/or modify
it under the terms of the GNU General Public License as published by
the Free Software Foundation; version 2 of the License.
🤝 مشارکت
اگر باگ یا پیشنهادی دارید، خوشحال می‌شوم از طریق بخش Issues یا Pull Request در GitHub اطلاع دهید.
