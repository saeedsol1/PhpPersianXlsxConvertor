<?php
// مبدل فایل با PhpSpreadsheet - سازگار با PHP 8.x
// ورودی: csv, txt, xlsx, xls, ods   |   خروجی: xlsx, xls, ods, csv, html, pdf
// فایل‌های تبدیل‌شده در پوشهٔ converted ذخیره و بعد از یک ساعت خودکار حذف می‌شوند.

date_default_timezone_set('Asia/Tehran');
header('Content-Type: text/html; charset=UTF-8');

// اگر پوشهٔ vendor جای دیگری است، مسیر را اصلاح کنید
require_once __DIR__ . '/vendor/autoload.php';

use PhpOffice\PhpSpreadsheet\IOFactory;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Reader\Csv as CsvReader;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx as XlsxWriter;
use PhpOffice\PhpSpreadsheet\Writer\Xls as XlsWriter;
use PhpOffice\PhpSpreadsheet\Writer\Ods as OdsWriter;
use PhpOffice\PhpSpreadsheet\Writer\Csv as CsvWriter;
use PhpOffice\PhpSpreadsheet\Writer\Html as HtmlWriter;
use PhpOffice\PhpSpreadsheet\Writer\Pdf\Mpdf as MpdfWriter;
use PhpOffice\PhpSpreadsheet\Worksheet\PageSetup;

@set_time_limit(300);
@ini_set('memory_limit', '512M');

const STORE_DIR = __DIR__ . '/converted';
const TTL       = 3600;   // مدت نگهداری فایل (ثانیه) = یک ساعت
const IN_EXT    = ['csv', 'txt', 'xlsx', 'xls', 'ods'];

/* ------------------------------------------------------------------ */
/* توابع کمکی                                                          */
/* ------------------------------------------------------------------ */

function h(string $s): string
{
	return htmlspecialchars($s, ENT_QUOTES, 'UTF-8');
}

/** PDF فقط وقتی فعال است که mpdf نصب شده باشد: composer require mpdf/mpdf */
function pdfAvailable(): bool
{
	return class_exists('\\Mpdf\\Mpdf');
}

/** نوع‌های خروجی: کلید => [عنوان، پسوند، MIME] */
function formats(): array
{
	return [
		'xlsx' => ['Excel (XLSX)',            'xlsx', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'],
		'xls'  => ['Excel 97-2003 (XLS)',     'xls',  'application/vnd.ms-excel'],
		'ods'  => ['OpenDocument (ODS)',      'ods',  'application/vnd.oasis.opendocument.spreadsheet'],
		'csv'  => ['CSV (UTF-8)',             'csv',  'text/csv; charset=UTF-8'],
		'html' => ['صفحه وب (HTML)',          'html', 'text/html; charset=UTF-8'],
		'pdf'  => ['PDF',                     'pdf',  'application/pdf'],
	];
}

function fa(string $s): string
{
	return strtr($s, ['0' => '۰', '1' => '۱', '2' => '۲', '3' => '۳', '4' => '۴',
	                  '5' => '۵', '6' => '۶', '7' => '۷', '8' => '۸', '9' => '۹']);
}

/** میلادی به شمسی */
function g2j(int $gy, int $gm, int $gd): array
{
	$gdm  = [0, 31, 59, 90, 120, 151, 181, 212, 243, 273, 304, 334];
	$gy2  = ($gm > 2) ? ($gy + 1) : $gy;
	$days = 355666 + (365 * $gy) + intdiv($gy2 + 3, 4) - intdiv($gy2 + 99, 100)
	      + intdiv($gy2 + 399, 400) + $gd + $gdm[$gm - 1];
	$jy   = -1595 + (33 * intdiv($days, 12053));
	$days %= 12053;
	$jy  += 4 * intdiv($days, 1461);
	$days %= 1461;
	if ($days > 365) {
		$jy  += intdiv($days - 1, 365);
		$days = ($days - 1) % 365;
	}
	if ($days < 186) {
		$jm = 1 + intdiv($days, 31);
		$jd = 1 + ($days % 31);
	} else {
		$jm = 7 + intdiv($days - 186, 30);
		$jd = 1 + (($days - 186) % 30);
	}
	return [$jy, $jm, $jd];
}

/** تاریخ و ساعت شمسی، مثلا ۱۴۰۵/۰۸/۲۷ - ۱۴:۳۰ */
function jdate(int $ts): string
{
	[$jy, $jm, $jd] = g2j((int)date('Y', $ts), (int)date('n', $ts), (int)date('j', $ts));
	return fa(sprintf('%04d/%02d/%02d', $jy, $jm, $jd) . ' - ' . date('H:i', $ts));
}

function fmtSize(int $b): string
{
	if ($b >= 1048576) { return fa(number_format($b / 1048576, 1)) . ' مگابایت'; }
	if ($b >= 1024)    { return fa(number_format($b / 1024, 1)) . ' کیلوبایت'; }
	return fa((string)$b) . ' بایت';
}

/** پوشهٔ ذخیره‌سازی را می‌سازد و دسترسی مستقیم از وب را (روی Apache) می‌بندد */
function ensureStore(): void
{
	if (!is_dir(STORE_DIR)) {
		mkdir(STORE_DIR, 0755, true);
	}
	$ht = STORE_DIR . '/.htaccess';
	if (!is_file($ht)) {
		file_put_contents($ht,
			"<IfModule mod_authz_core.c>\nRequire all denied\n</IfModule>\n" .
			"<IfModule !mod_authz_core.c>\nOrder allow,deny\nDeny from all\n</IfModule>\n");
	}
}

/** هر فایلی که از ساختنش بیش از یک ساعت گذشته حذف می‌شود */
function cleanup(): void
{
	$files = glob(STORE_DIR . '/*');
	foreach ($files ?: [] as $f) {
		if (is_file($f) && filemtime($f) + TTL <= time()) {
			@unlink($f);
		}
	}
}

function loadMeta(string $id): ?array
{
	if (!preg_match('/^[a-f0-9]{16}$/', $id)) { return null; }
	$j = STORE_DIR . '/' . $id . '.json';
	if (!is_file($j)) { return null; }
	$m = json_decode((string)file_get_contents($j), true);
	if (!is_array($m) || !isset($m['id'], $m['file'], $m['expires'], $m['created'])) { return null; }
	if ($m['expires'] <= time() || !is_file(STORE_DIR . '/' . $m['file'])) { return null; }
	return $m;
}

function listFiles(): array
{
	$out   = [];
	$files = glob(STORE_DIR . '/*.json');
	foreach ($files ?: [] as $j) {
		$m = loadMeta(basename($j, '.json'));
		if ($m) { $out[] = $m; }
	}
	usort($out, fn($a, $b) => $b['created'] <=> $a['created']);
	return $out;
}

/** تشخیص جداکننده از روی اولین خط */
function detectDelimiter(string $path): string
{
	$fh   = fopen($path, 'rb');
	$line = $fh ? (string)fgets($fh, 65536) : '';
	if ($fh) { fclose($fh); }
	if (strncmp($line, "\xEF\xBB\xBF", 3) === 0) { $line = substr($line, 3); }

	$best = ',';
	$max  = 0;
	foreach ([',', ';', "\t", '|'] as $d) {
		$c = substr_count($line, $d);
		if ($c > $max) { $max = $c; $best = $d; }
	}
	return $best;
}

/** تشخیص کدگذاری: UTF-8 (با/بدون BOM)، UTF-16 یا Windows-1256 */
function detectEncoding(string $path): string
{
	$fh     = fopen($path, 'rb');
	$sample = $fh ? (string)fread($fh, 200000) : '';
	if ($fh) { fclose($fh); }

	if (strncmp($sample, "\xEF\xBB\xBF", 3) === 0) { return 'UTF-8'; }
	if (strncmp($sample, "\xFF\xFE", 2) === 0)     { return 'UTF-16LE'; }
	if (strncmp($sample, "\xFE\xFF", 2) === 0)     { return 'UTF-16BE'; }

	$check = strlen($sample) > 3 ? substr($sample, 0, -3) : $sample;
	return mb_check_encoding($check, 'UTF-8') ? 'UTF-8' : 'Windows-1256';
}

function makeWriter(Spreadsheet $ss, string $fmt)
{
	switch ($fmt) {
		case 'xlsx': return new XlsxWriter($ss);
		case 'xls':  return new XlsWriter($ss);
		case 'ods':  return new OdsWriter($ss);
		case 'csv':
			$w = new CsvWriter($ss);
			$w->setUseBOM(true);          // تا اکسل فارسی را درست بخواند
			$w->setDelimiter(',');
			$w->setEnclosure('"');
			return $w;
		case 'html': return new HtmlWriter($ss);
		case 'pdf':  return new MpdfWriter($ss);
	}
	throw new Exception('نوع خروجی نامعتبر است.');
}

/** تبدیل فایل و ذخیره در پوشهٔ converted؛ متادیتا برمی‌گرداند */
function convertFile(string $in, string $origName, string $target): array
{
	$formats = formats();
	$origName = preg_replace('~.*[\\\\/]~', '', $origName);
	$ext = strtolower(pathinfo($origName, PATHINFO_EXTENSION));

	if (in_array($ext, ['csv', 'txt'], true)) {
		$reader = new CsvReader();
		$reader->setInputEncoding(detectEncoding($in));
		$reader->setDelimiter(detectDelimiter($in));
		$reader->setEnclosure('"');
		$reader->setSheetIndex(0);
		$ss    = $reader->load($in);
		$sheet = $ss->getActiveSheet();

		// شیت راست‌به‌چپ، هدر پررنگ و ثابت، عرض ستون خودکار
		$sheet->setRightToLeft(true);
		$sheet->getStyle('A1:' . $sheet->getHighestColumn() . '1')->getFont()->setBold(true);
		$sheet->freezePane('A2');
		if ($sheet->getHighestRow() <= 5000) {
			foreach ($sheet->getColumnIterator() as $col) {
				$sheet->getColumnDimension($col->getColumnIndex())->setAutoSize(true);
			}
		}
	} else {
		$ss = IOFactory::load($in);   // xlsx / xls / ods
	}

	if ($target === 'pdf') {
		$ss->getDefaultStyle()->getFont()->setName('DejaVu Sans')->setSize(7);
		$ps = $ss->getActiveSheet()->getPageSetup();
		$ps->setOrientation(PageSetup::ORIENTATION_LANDSCAPE);
		$ps->setPaperSize(PageSetup::PAPERSIZE_A4);
	}

	ensureStore();
	$id   = bin2hex(random_bytes(8));
	$file = $id . '.' . $formats[$target][1];
	makeWriter($ss, $target)->save(STORE_DIR . '/' . $file);
	$ss->disconnectWorksheets();

	$now  = time();
	$meta = [
		'id'      => $id,
		'name'    => $origName,
		'from'    => $ext,
		'format'  => $target,
		'file'    => $file,
		'size'    => (int)filesize(STORE_DIR . '/' . $file),
		'created' => $now,
		'expires' => $now + TTL,
	];
	file_put_contents(STORE_DIR . '/' . $id . '.json', json_encode($meta, JSON_UNESCAPED_UNICODE));
	return $meta;
}

/* ------------------------------------------------------------------ */
/* منطق صفحه                                                           */
/* ------------------------------------------------------------------ */

ensureStore();
cleanup();
$formats = formats();

// دانلود
if (isset($_GET['download'])) {
	$m = loadMeta((string)$_GET['download']);
	if (!$m) {
		http_response_code(404);
		echo 'فایل پیدا نشد یا منقضی شده است.';
		exit;
	}
	$f    = $formats[$m['format']] ?? null;
	$ext  = $f ? $f[1] : pathinfo($m['file'], PATHINFO_EXTENSION);
	$mime = $f ? $f[2] : 'application/octet-stream';
	$base = preg_replace('/[^\w\-]+/u', '_', pathinfo($m['name'], PATHINFO_FILENAME));
	$name = ($base !== '' ? $base : 'myexcelfile') . '.' . $ext;

	while (ob_get_level() > 0) { ob_end_clean(); }
	header('Content-Type: ' . $mime);
	header('Content-Disposition: attachment; filename="download.' . $ext . '"; filename*=UTF-8\'\'' . rawurlencode($name));
	header('Content-Length: ' . filesize(STORE_DIR . '/' . $m['file']));
	header('Cache-Control: max-age=0');
	readfile(STORE_DIR . '/' . $m['file']);
	exit;
}

$modal = null;   // ['type' => 'done'|'error', ...]

if ($_SERVER['REQUEST_METHOD'] === 'POST' && isset($_FILES['file'])) {
	try {
		if ($_FILES['file']['error'] !== UPLOAD_ERR_OK) {
			throw new Exception('خطا در آپلود فایل (کد ' . (int)$_FILES['file']['error'] . ')');
		}
		$target = (string)($_POST['format'] ?? 'xlsx');
		if (!isset($formats[$target])) {
			throw new Exception('نوع خروجی نامعتبر است.');
		}
		if ($target === 'pdf' && !pdfAvailable()) {
			throw new Exception('برای خروجی PDF باید ابتدا دستور composer require mpdf/mpdf اجرا شود.');
		}
		$ext = strtolower(pathinfo((string)$_FILES['file']['name'], PATHINFO_EXTENSION));
		if (!in_array($ext, IN_EXT, true)) {
			throw new Exception('فقط فایل‌های ' . implode('، ', IN_EXT) . ' مجاز است.');
		}

		$meta = convertFile($_FILES['file']['tmp_name'], (string)$_FILES['file']['name'], $target);
		header('Location: index.php?done=' . $meta['id']);   // جلوگیری از ارسال مجدد فرم با Refresh
		exit;
	} catch (Throwable $e) {
		$modal = ['type' => 'error', 'msg' => $e->getMessage()];
	}
} elseif (isset($_GET['done'])) {
	$m = loadMeta((string)$_GET['done']);
	if ($m) { $modal = ['type' => 'done', 'meta' => $m]; }
}

$files = listFiles();
?>
<!DOCTYPE html>
<html lang="fa" dir="rtl">
<head>
	<meta charset="utf-8">
	<meta name="viewport" content="width=device-width, initial-scale=1">
	<title>تبدیل فایل</title>
<link rel="stylesheet" href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.9/dist/css/bootstrap.rtl.min.css" integrity="sha384-fFMX5LqaHoT6olm8r/JTe3Dg4+vFEFXrVlhIo8IGs96iA+JhgKLLYwY4y2fHeyMY" crossorigin="anonymous">
	<style>
		/* مودال با تگ استاندارد dialog (بدون نیاز به bootstrap.js) */
		dialog.rm { border: 0; border-radius: 12px; padding: 0; width: 92%; max-width: 480px; box-shadow: 0 10px 40px rgba(0,0,0,.3); }
		dialog.rm::backdrop { background: rgba(0,0,0,.5); }
		.rm-head { padding: 14px 18px; font-weight: bold; color: #fff; }
		.rm-head.done { background: #198754; }
		.rm-head.error { background: #dc3545; }
		.rm-body { padding: 18px; line-height: 2; }
		.rm-foot { padding: 12px 18px; display: flex; gap: 8px; justify-content: flex-start; border-top: 1px solid #dee2e6; }
		.rm-foot form { margin: 0; }
		.files-wrap { max-width: 1000px; margin: 30px auto; padding: 0 12px; overflow-x: auto; }
		.files-wrap td, .files-wrap th { vertical-align: middle; white-space: nowrap; }
	</style>
</head>

<body>

	<h3 class="d-flex justify-content-center p-2">Convert CSV and TXT to XLSX</h3><br>

	<h2 class="d-flex justify-content-center">تبدیل فایلهای متنی به اکسل</h2><br>

	<div class="d-flex justify-content-center">
		<form action="index.php" method="post" enctype="multipart/form-data" class="d-flex flex-wrap gap-2 align-items-center justify-content-center">
			<input class="form-control" style="width:auto" type="file" name="file" accept=".csv,.txt,.xlsx,.xls,.ods" required>
			<label for="format" class="mb-0">تبدیل به:</label>
			<select class="form-select" style="width:auto" id="format" name="format">
				<?php foreach ($formats as $key => $f):
					$disabled = ($key === 'pdf' && !pdfAvailable()); ?>
					<option value="<?= h($key) ?>"<?= $disabled ? ' disabled' : '' ?>>
						<?= h($f[0]) ?><?= $disabled ? ' (نیاز به mpdf)' : '' ?>
					</option>
				<?php endforeach; ?>
			</select>
			<button type="submit" class="btn btn-danger">تبدیل</button>
		</form>
	</div>
	<p class="text-center text-muted mt-2" style="font-size:.9rem">
		ورودی: CSV، TXT، XLSX، XLS، ODS &nbsp;|&nbsp; فایل‌های تبدیل‌شده یک ساعت نگهداری و سپس خودکار حذف می‌شوند.
	</p>

	<div class="files-wrap">
		<h5>فایل‌های آماده دانلود</h5>
		<?php if (!$files): ?>
			<p class="text-muted">فعلا فایلی وجود ندارد.</p>
		<?php else: ?>
		<table class="table table-striped table-sm">
			<thead>
				<tr>
					<th>#</th>
					<th>نام فایل</th>
					<th>نوع تبدیل</th>
					<th>حجم</th>
					<th>تاریخ ایجاد</th>
					<th>تاریخ انقضا</th>
					<th>زمان باقی‌مانده</th>
					<th>دانلود</th>
				</tr>
			</thead>
			<tbody>
			<?php foreach ($files as $i => $m): ?>
				<tr>
					<td><?= fa((string)($i + 1)) ?></td>
					<td><?= h($m['name']) ?></td>
					<td dir="ltr" style="text-align:right"><?= h(strtoupper($m['from'])) ?> → <?= h(strtoupper($m['format'])) ?></td>
					<td><?= h(fmtSize((int)$m['size'])) ?></td>
					<td><?= h(jdate((int)$m['created'])) ?></td>
					<td><?= h(jdate((int)$m['expires'])) ?></td>
					<td data-left="<?= max(0, (int)$m['expires'] - time()) ?>">--:--</td>
					<td><a class="btn btn-sm btn-success" href="index.php?download=<?= h($m['id']) ?>">دانلود</a></td>
				</tr>
			<?php endforeach; ?>
			</tbody>
		</table>
		<?php endif; ?>
	</div>

<?php if ($modal): ?>
	<dialog id="resultModal" class="rm">
	<?php if ($modal['type'] === 'done'):
		$m = $modal['meta']; ?>
		<div class="rm-head done">تبدیل با موفقیت انجام شد</div>
		<div class="rm-body">
			<div>نام فایل: <b><?= h($m['name']) ?></b></div>
			<div>نوع تبدیل: <b dir="ltr"><?= h(strtoupper($m['from'])) ?> → <?= h(strtoupper($m['format'])) ?></b></div>
			<div>حجم: <?= h(fmtSize((int)$m['size'])) ?></div>
			<div>این فایل تا <b><?= h(jdate((int)$m['expires'])) ?></b> (یک ساعت) قابل دانلود است و بعد از آن حذف می‌شود.</div>
		</div>
		<div class="rm-foot">
			<a class="btn btn-success" href="index.php?download=<?= h($m['id']) ?>">دانلود فایل</a>
			<form method="dialog"><button class="btn btn-secondary">بستن</button></form>
		</div>
	<?php else: ?>
		<div class="rm-head error">خطا در تبدیل</div>
		<div class="rm-body"><?= h($modal['msg']) ?></div>
		<div class="rm-foot">
			<form method="dialog"><button class="btn btn-secondary">بستن</button></form>
		</div>
	<?php endif; ?>
	</dialog>
	<script>
		(function () {
			var d = document.getElementById('resultModal');
			if (!d || !d.showModal) { return; }
			d.showModal();
			d.addEventListener('close', function () {
				if (location.search) { history.replaceState(null, '', 'index.php'); }
			});
		})();
	</script>
<?php endif; ?>

	<script>
		// شمارش معکوس زمان باقی‌مانده هر فایل
		(function () {
			var cells = document.querySelectorAll('[data-left]');
			if (!cells.length) { return; }
			var start = Date.now(), reloading = false;
			var digits = '۰۱۲۳۴۵۶۷۸۹';
			function fa(s) { return s.replace(/\d/g, function (d) { return digits.charAt(d); }); }
			function tick() {
				var elapsed = (Date.now() - start) / 1000, expired = false;
				cells.forEach(function (c) {
					var left = Math.floor(parseInt(c.getAttribute('data-left'), 10) - elapsed);
					if (left <= 0) { expired = true; c.textContent = fa('00:00'); return; }
					var m = Math.floor(left / 60), s = left % 60;
					c.textContent = fa((m < 10 ? '0' : '') + m + ':' + (s < 10 ? '0' : '') + s);
				});
				if (expired && !reloading) {
					reloading = true;
					setTimeout(function () { location.href = 'index.php'; }, 1500);
				}
			}
			tick();
			setInterval(tick, 1000);
		})();
	</script>
</body>
</html>
