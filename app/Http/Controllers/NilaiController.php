<?php



namespace App\Http\Controllers;



use App\Models\Nilai;

use App\Models\MataKuliah;

use Illuminate\Http\Request;

use Illuminate\Support\Facades\Validator;

use PhpOffice\PhpSpreadsheet\IOFactory;

use PhpOffice\PhpSpreadsheet\Spreadsheet;

use PhpOffice\PhpSpreadsheet\Writer\Xlsx;

use FPDF;



class NilaiController extends Controller

{

    /*

    |--------------------------------------------------------------------------

    | HALAMAN NILAI AKADEMIK

    |--------------------------------------------------------------------------

    */



    public function index(Request $request)

    {

        $sort = $request->get('sort', 'created_at');

        $direction = $request->get('direction', 'desc');



        $allowedSorts = [

            'created_at',

            'nilai',

            'jenis_penilaian',

            'mata_kuliah',

        ];



        if (!in_array($sort, $allowedSorts)) {

            $sort = 'created_at';

        }



        if (!in_array($direction, ['asc', 'desc'])) {

            $direction = 'desc';

        }



        $query = Nilai::with('mataKuliah');



        if ($sort === 'mata_kuliah') {

            $query->orderBy(

                MataKuliah::select('nama')

                    ->whereColumn(

                        'mata_kuliah.id',

                        'nilais.mata_kuliah_id'

                    ),

                $direction

            );

        } else {

            $query->orderBy($sort, $direction);

        }



        $nilai = $query->paginate(10)->withQueryString();



        $rataRata = Nilai::avg('nilai');



        $totalMataKuliah = Nilai::distinct('mata_kuliah_id')

            ->count('mata_kuliah_id');



        $daftarSemester = MataKuliah::select('semester')

            ->distinct()

            ->orderBy('semester')

            ->pluck('semester');



        return view('nilai.index', compact(

            'nilai',

            'rataRata',

            'totalMataKuliah',

            'sort',

            'direction',

            'daftarSemester'

        ));

    }



    /*

    |--------------------------------------------------------------------------

    | CRUD NILAI

    |--------------------------------------------------------------------------

    */



    public function create()

    {

        $mataKuliah = MataKuliah::orderBy('nama')->get();



        return view('nilai.create', compact('mataKuliah'));

    }



    public function store(Request $request)

    {

        $data = $request->validate([

            'mata_kuliah_id' => 'required|exists:mata_kuliah,id',

            'jenis_penilaian' => 'required|string|max:100',

            'nilai' => 'required|numeric|min:0|max:100',

            'keterangan' => 'nullable|string',

        ]);



        $data['keterangan'] = $data['keterangan'] ?: null;



        if ($this->isDuplicate($data)) {

            return back()->withInput()->withErrors([

                'nilai' => 'Catatan nilai dengan data yang sama sudah tersedia.',

            ]);

        }



        Nilai::create($data);



        return redirect()->route('nilai.index')

            ->with('success', 'Nilai berhasil ditambahkan.');

    }



    public function edit(Nilai $nilai)

    {

        $mataKuliah = MataKuliah::orderBy('nama')->get();



        return view('nilai.edit', compact('nilai', 'mataKuliah'));

    }



    public function update(Request $request, Nilai $nilai)

    {

        $data = $request->validate([

            'mata_kuliah_id' => 'required|exists:mata_kuliah,id',

            'jenis_penilaian' => 'required|string|max:100',

            'nilai' => 'required|numeric|min:0|max:100',

            'keterangan' => 'nullable|string',

        ]);



        $data['keterangan'] = $data['keterangan'] ?: null;



        if ($this->isDuplicate($data, $nilai->id)) {

            return back()->withInput()->withErrors([

                'nilai' => 'Catatan nilai dengan data yang sama sudah tersedia.',

            ]);

        }



        $nilai->update($data);



        return redirect()->route('nilai.index')

            ->with('success', 'Nilai berhasil diperbarui.');

    }



    public function destroy(Nilai $nilai)

    {

        $nilai->delete();



        return redirect()->route('nilai.index')

            ->with('success', 'Nilai berhasil dihapus.');

    }



    private function isDuplicate(array $data, ?int $exceptId = null): bool

    {

        $query = Nilai::where('mata_kuliah_id', $data['mata_kuliah_id'])

            ->where('jenis_penilaian', $data['jenis_penilaian'])

            ->where('nilai', $data['nilai']);



        if ($data['keterangan'] === null) {

            $query->whereNull('keterangan');

        } else {

            $query->where('keterangan', $data['keterangan']);

        }



        if ($exceptId !== null) {

            $query->where('id', '!=', $exceptId);

        }



        return $query->exists();

    }



    /*

    |--------------------------------------------------------------------------

    | EXPORT EXCEL

    |--------------------------------------------------------------------------

    */



    public function export()

    {

        $data = Nilai::with('mataKuliah')

            ->orderBy('created_at', 'desc')

            ->get();



        $spreadsheet = new Spreadsheet();

        $sheet = $spreadsheet->getActiveSheet();

        $sheet->setTitle('Data Nilai');



        $sheet->fromArray([

            ['Kode MK', 'Mata Kuliah', 'Jenis Penilaian', 'Nilai', 'Keterangan']

        ], null, 'A1');



        $row = 2;



        foreach ($data as $item) {

            $sheet->fromArray([[

                $item->mataKuliah->kode ?? '-',

                $item->mataKuliah->nama ?? '-',

                $item->jenis_penilaian,

                (float) $item->nilai,

                $item->keterangan,

            ]], null, 'A' . $row);



            $row++;

        }



        $sheet->getStyle('A1:E1')->getFont()->setBold(true);



        foreach (range('A', 'E') as $column) {

            $sheet->getColumnDimension($column)->setAutoSize(true);

        }



        $info = $spreadsheet->createSheet();

        $info->setTitle('Informasi Export');

        $info->setCellValue('A1', 'Waktu Export');

        $info->setCellValue('B1', now()->format('d-m-Y H:i:s'));

        $info->setCellValue('A2', 'Total Data');

        $info->setCellValue('B2', $data->count());



        $filename = 'StudentTrack_Nilai_' . now()->format('Ymd_His') . '.xlsx';



        $writer = new Xlsx($spreadsheet);



        return response()->streamDownload(function () use ($writer) {

            $writer->save('php://output');

        }, $filename, [

            'Content-Type' =>

                'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',

        ]);

    }



    /*

    |--------------------------------------------------------------------------

    | IMPORT EXCEL

    |--------------------------------------------------------------------------

    */



    public function import(Request $request)

    {

        $request->validate([

            'file' => 'required|file|mimes:xlsx,xls,csv|max:5120',

        ]);



        try {

            $spreadsheet = IOFactory::load(

                $request->file('file')->getRealPath()

            );



            $sheet = $spreadsheet->getSheetByName('Data Nilai')

                ?? $spreadsheet->getActiveSheet();



            $rows = $sheet->toArray();



            $berhasil = 0;

            $duplikat = 0;

            $gagal = 0;



            foreach (array_slice($rows, 1) as $row) {

                $kode = trim((string) ($row[0] ?? ''));

                $jenis = trim((string) ($row[2] ?? ''));

                $nilai = $row[3] ?? null;

                $keterangan = trim((string) ($row[4] ?? ''));



                if ($kode === '' && $jenis === '' && $nilai === null) {

                    continue;

                }



                $mataKuliah = MataKuliah::where('kode', $kode)->first();



                $validator = Validator::make([

                    'jenis_penilaian' => $jenis,

                    'nilai' => $nilai,

                ], [

                    'jenis_penilaian' => 'required|string|max:100',

                    'nilai' => 'required|numeric|min:0|max:100',

                ]);



                if (!$mataKuliah || $validator->fails()) {

                    $gagal++;

                    continue;

                }



                $data = [

                    'mata_kuliah_id' => $mataKuliah->id,

                    'jenis_penilaian' => $jenis,

                    'nilai' => $nilai,

                    'keterangan' => $keterangan !== '' ? $keterangan : null,

                ];



                if ($this->isDuplicate($data)) {

                    $duplikat++;

                    continue;

                }



                Nilai::create($data);

                $berhasil++;

            }



            return redirect()->route('nilai.index')->with(

                'success',

                "Import selesai. Berhasil: {$berhasil}, duplikat: {$duplikat}, tidak valid: {$gagal}."

            );



        } catch (\Throwable $e) {

            return back()->withErrors([

                'file' => 'File gagal dibaca: ' . $e->getMessage(),

            ]);

        }

    }



    /*

    |--------------------------------------------------------------------------

    | NORMALISASI KOMPONEN PENILAIAN

    |--------------------------------------------------------------------------

    */



    private function normalisasiKomponen(string $jenis): ?string

    {

        $jenis = strtolower(trim($jenis));

        $jenis = preg_replace('/\s+/', ' ', $jenis);



        $komponen = [

            'project_akhir' => [

                'project akhir',

                'projek akhir',

                'project',

                'proyek akhir',

            ],



            'tugas' => [

                'tugas',

                'assignment',

            ],



            'uts' => [

                'uts',

                'ujian tengah semester',

            ],



            'uas' => [

                'uas',

                'ujian akhir semester',

            ],



            'keaktifan' => [

                'keaktifan',

                'aktivitas',

                'partisipasi',

            ],



            'quiz' => [

                'quiz',

                'kuis',

            ],

        ];



        foreach ($komponen as $nama => $alias) {

            if (in_array($jenis, $alias, true)) {

                return $nama;

            }

        }



        return null;

    }



    /*

    |--------------------------------------------------------------------------

    | BOBOT PENILAIAN

    |--------------------------------------------------------------------------

    */



    private function bobotKomponen(): array

    {

        return [

            'project_akhir' => 0.50,

            'tugas' => 0.20,

            'uts' => 0.10,

            'uas' => 0.10,

            'keaktifan' => 0.05,

            'quiz' => 0.05,

        ];

    }



    /*

    |--------------------------------------------------------------------------

    | KONVERSI NILAI ANGKA KE HURUF

    |--------------------------------------------------------------------------

    */



    private function konversiNilai(float $nilai): array

    {

        if ($nilai >= 85) {

            return ['A', 4.00];

        }



        if ($nilai >= 80) {

            return ['AB', 3.50];

        }



        if ($nilai >= 75) {

            return ['B', 3.00];

        }



        if ($nilai >= 70) {

            return ['BC', 2.50];

        }



        if ($nilai >= 60) {

            return ['C', 2.00];

        }



        if ($nilai >= 50) {

            return ['D', 1.00];

        }



        return ['E', 0.00];

    }



    /*

    |--------------------------------------------------------------------------

    | PERHITUNGAN NILAI AKHIR PER MATA KULIAH

    |--------------------------------------------------------------------------

    */



    private function dataAkademik(?int $semester = null): array

    {

        $query = MataKuliah::with('nilai');



        if ($semester !== null) {

            $query->where('semester', $semester);

        }



        $mataKuliah = $query

            ->orderBy('semester')

            ->orderBy('nama')

            ->get();



        $bobotKomponen = $this->bobotKomponen();



        $hasil = [];



        foreach ($mataKuliah as $mk) {

            $komponen = array_fill_keys(

                array_keys($bobotKomponen),

                []

            );



            foreach ($mk->nilai as $item) {

                $namaKomponen = $this->normalisasiKomponen(

                    $item->jenis_penilaian

                );



                if ($namaKomponen !== null) {

                    $komponen[$namaKomponen][] = (float) $item->nilai;

                }

            }



            // Menghitung rata-rata setiap komponen.

            $rataKomponen = [];



            foreach ($komponen as $nama => $daftarNilai) {

                $rataKomponen[$nama] = count($daftarNilai) > 0

                    ? array_sum($daftarNilai) / count($daftarNilai)

                    : null;

            }



            // Hitung nilai akhir dari komponen yang tersedia saja.
            // Bobot komponen yang tersedia dinormalisasi agar totalnya menjadi 100%.
            $totalBobotTersedia = 0.0;
            $totalNilaiTertimbang = 0.0;

            foreach ($bobotKomponen as $nama => $persentase) {
                if ($rataKomponen[$nama] !== null) {
                    $totalNilaiTertimbang +=
                        $rataKomponen[$nama] * $persentase;
                    $totalBobotTersedia += $persentase;
                }
            }

            // Minimal satu komponen nilai harus tersedia agar dapat dihitung.
            $adaNilai = $totalBobotTersedia > 0;
            $nilaiAkhir = null;
            $huruf = '-';
            $bobot = null;

            if ($adaNilai) {
                $nilaiAkhir = $totalNilaiTertimbang / $totalBobotTersedia;
                [$huruf, $bobot] = $this->konversiNilai($nilaiAkhir);
            }


            $hasil[] = [

                'id' => $mk->id,

                'kode' => $mk->kode,

                'nama' => $mk->nama,

                'semester' => $mk->semester,

                'sks' => (int) $mk->sks,

                'komponen' => $rataKomponen,

                'lengkap' => $adaNilai,

                'nilai_akhir' => $nilaiAkhir,

                'huruf' => $huruf,

                'bobot' => $bobot,

            ];

        }



        return $hasil;

    }



    /*

    |--------------------------------------------------------------------------

    | PERHITUNGAN IPS DAN IPK

    |--------------------------------------------------------------------------

    */



    private function hitungIndeks(array $data): array

    {

        $totalSks = 0;

        $totalMutu = 0;

        $totalSksSemua = 0;



        foreach ($data as $item) {

            $totalSksSemua += $item['sks'];



            if ($item['lengkap']) {

                $totalSks += $item['sks'];

                $totalMutu += $item['sks'] * $item['bobot'];

            }

        }



        return [

            'ips' => $totalSks > 0

                ? $totalMutu / $totalSks

                : 0,



            'sks_dihitung' => $totalSks,

            'total_sks' => $totalSksSemua,

            'total_mutu' => $totalMutu,

        ];

    }



    /*

    |--------------------------------------------------------------------------

    | DOWNLOAD KHS PDF

    |--------------------------------------------------------------------------

    */



    public function downloadKhs($semester)

    {

        $semester = (int) $semester;



        if ($semester < 1 || $semester > 20) {

            abort(404, 'Semester tidak valid.');

        }



        $data = $this->dataAkademik($semester);



        if (count($data) === 0) {

            return redirect()->route('nilai.index')

                ->with(

                    'error',

                    'Tidak ada mata kuliah pada semester tersebut.'

                );

        }



        $indeks = $this->hitungIndeks($data);



        $pdf = new FPDF('L', 'mm', 'A4');

        $pdf->AddPage();

        $pdf->SetAutoPageBreak(true, 15);



        $this->headerPdf(

            $pdf,

            'KARTU HASIL STUDI',

            'Semester ' . $semester

        );



        $this->tabelNilai($pdf, $data);



        $pdf->Ln(5);

        $pdf->SetFont('Arial', 'B', 11);

        $pdf->Cell(0, 8, 'RINGKASAN HASIL STUDI', 0, 1);



        $pdf->SetFont('Arial', '', 10);



        $pdf->Cell(

            60,

            8,

            'Total SKS: ' . $indeks['total_sks'],

            0,

            0

        );



        $pdf->Cell(

            70,

            8,

            'SKS Dihitung: ' . $indeks['sks_dihitung'],

            0,

            0

        );



        $pdf->Cell(

            60,

            8,

            'IPS: ' . number_format($indeks['ips'], 2),

            0,

            1

        );



        $pdf->Ln(5);

        $pdf->SetFont('Arial', 'I', 8);



        $pdf->MultiCell(

            0,

            5,

            $this->pdfText(

                'Catatan: Nilai akhir dihitung dari komponen yang tersedia; bobot komponen yang tersedia dinormalisasi menjadi 100%. Mata kuliah tanpa nilai tidak dihitung. Dokumen ini merupakan transkrip simulasi StudentTrack, bukan dokumen akademik resmi.'

            )

        );



        return $this->kirimPdf(

            $pdf,

            'StudentTrack_KHS_Semester_' . $semester . '.pdf'

        );

    }



    /*

    |--------------------------------------------------------------------------

    | DOWNLOAD TRANSKRIP PDF

    |--------------------------------------------------------------------------

    */



    public function downloadTranskrip()

    {

        $data = $this->dataAkademik();



        if (count($data) === 0) {

            return redirect()->route('nilai.index')

                ->with('error', 'Belum ada data mata kuliah.');

        }



        $indeks = $this->hitungIndeks($data);



        $pdf = new FPDF('L', 'mm', 'A4');

        $pdf->AddPage();

        $pdf->SetAutoPageBreak(true, 15);



        $this->headerPdf(

            $pdf,

            'TRANSKRIP NILAI AKADEMIK',

            'Seluruh Semester'

        );



        $this->tabelNilai($pdf, $data);



        $pdf->Ln(5);

        $pdf->SetFont('Arial', 'B', 11);

        $pdf->Cell(0, 8, 'RINGKASAN TRANSKRIP', 0, 1);



        $pdf->SetFont('Arial', '', 10);



        $pdf->Cell(

            60,

            8,

            'Total SKS: ' . $indeks['total_sks'],

            0,

            0

        );



        $pdf->Cell(

            70,

            8,

            'SKS Dihitung: ' . $indeks['sks_dihitung'],

            0,

            0

        );



        $pdf->Cell(

            60,

            8,

            'IPK: ' . number_format($indeks['ips'], 2),

            0,

            1

        );



        $pdf->Ln(5);

        $pdf->SetFont('Arial', 'I', 8);



        $pdf->MultiCell(

            0,

            5,

            $this->pdfText(

                'Catatan: IPK dihitung dari mata kuliah yang memiliki minimal satu komponen nilai. Bobot komponen yang tersedia dinormalisasi menjadi 100%. Dokumen ini merupakan transkrip simulasi StudentTrack, bukan dokumen akademik resmi.'

            )

        );



        return $this->kirimPdf(

            $pdf,

            'StudentTrack_Transkrip_Nilai.pdf'

        );

    }



    /*

    |--------------------------------------------------------------------------

    | HEADER PDF

    |--------------------------------------------------------------------------

    */



    private function headerPdf(

        FPDF $pdf,

        string $judul,

        string $subjudul

    ): void {

        $logo = public_path('images/logo-ipb.png');



        if (file_exists($logo)) {

            $pdf->Image($logo, 12, 10, 22);

        }



        $pdf->SetFont('Arial', 'B', 18);

        $pdf->SetTextColor(76, 61, 170);

        $pdf->Cell(0, 10, 'STUDENT TRACK', 0, 1, 'C');



        $pdf->SetFont('Arial', 'B', 14);

        $pdf->SetTextColor(35, 39, 60);

        $pdf->Cell(

            0,

            8,

            $this->pdfText($judul),

            0,

            1,

            'C'

        );



        $pdf->SetFont('Arial', '', 10);

        $pdf->SetTextColor(100, 100, 100);

        $pdf->Cell(

            0,

            7,

            $this->pdfText($subjudul),

            0,

            1,

            'C'

        );



        $pdf->Ln(5);



        $pdf->SetDrawColor(100, 90, 190);

        $pdf->Line(10, $pdf->GetY(), 287, $pdf->GetY());



        $pdf->Ln(5);

        $pdf->SetTextColor(40, 40, 40);

    }



    /*

    |--------------------------------------------------------------------------

    | TABEL NILAI PDF

    |--------------------------------------------------------------------------

    */



    private function tabelNilai(FPDF $pdf, array $data): void

    {

        $lebar = [

            10, 12, 42, 9,

            17, 14, 12, 12, 16, 16,

            17, 13, 13, 22

        ];



        $header = [

            'No',

            'Sem',

            'Mata Kuliah',

            'SKS',

            'Project Akhir',

            'Tugas',

            'UTS',

            'UAS',

            'Keaktifan',

            'Quiz',

            'Nilai Akhir',

            'Huruf',

            'Bobot',

            'Status'

        ];



        $gambarHeader = function () use ($pdf, $header, $lebar) {

            $pdf->SetFont('Arial', 'B', 7);

            $pdf->SetFillColor(235, 232, 250);



            foreach ($header as $i => $judul) {

                $pdf->Cell(

                    $lebar[$i],

                    10,

                    $judul,

                    1,

                    0,

                    'C',

                    true

                );

            }



            $pdf->Ln();

            $pdf->SetFont('Arial', '', 7);

        };



        $gambarHeader();



        foreach ($data as $index => $item) {

            if ($pdf->GetY() > 175) {

                $pdf->AddPage();



                $this->headerPdf(

                    $pdf,

                    'LANJUTAN TRANSKRIP NILAI',

                    'StudentTrack'

                );



                $gambarHeader();

            }



            $komponen = $item['komponen'];



            $baris = [

                $index + 1,

                $item['semester'],

                $this->pdfText($item['nama']),

                $item['sks'],



                $this->formatAngka(

                    $komponen['project_akhir']

                ),



                $this->formatAngka(

                    $komponen['tugas']

                ),



                $this->formatAngka(

                    $komponen['uts']

                ),



                $this->formatAngka(

                    $komponen['uas']

                ),



                $this->formatAngka(

                    $komponen['keaktifan']

                ),



                $this->formatAngka(

                    $komponen['quiz']

                ),



                $item['nilai_akhir'] !== null

                    ? number_format($item['nilai_akhir'], 2)

                    : '-',



                $item['huruf'],



                $item['bobot'] !== null

                    ? number_format($item['bobot'], 2)

                    : '-',



                $item['lengkap']

                    ? 'Lengkap'

                    : 'Belum lengkap',

            ];



            foreach ($baris as $i => $nilai) {

                $pdf->Cell(

                    $lebar[$i],

                    9,

                    (string) $nilai,

                    1,

                    0,

                    $i === 2 ? 'L' : 'C'

                );

            }



            $pdf->Ln();

        }

    }



    /*

    |--------------------------------------------------------------------------

    | FORMAT ANGKA

    |--------------------------------------------------------------------------

    */



    private function formatAngka(?float $nilai): string

    {

        return $nilai !== null

            ? number_format($nilai, 2)

            : '-';

    }



    /*

    |--------------------------------------------------------------------------

    | KONVERSI TEKS UNTUK FPDF

    |--------------------------------------------------------------------------

    */



    private function pdfText(string $text): string

    {

        // FPDF core fonts use a single-byte encoding. Avoid iconv TRANSLIT
        // because some hosting environments disable that conversion mode.
        if (function_exists('mb_convert_encoding')) {
            return mb_convert_encoding($text, 'ISO-8859-1', 'UTF-8');
        }

        // Fallback when mbstring is unavailable.
        $converted = @iconv('UTF-8', 'ISO-8859-1//IGNORE', $text);
        return $converted !== false ? $converted : '';

    }



    /*

    |--------------------------------------------------------------------------

    | MENGIRIM FILE PDF

    |--------------------------------------------------------------------------

    */



    private function kirimPdf(FPDF $pdf, string $filename)

    {

        $content = $pdf->Output('S');



        return response($content, 200, [

            'Content-Type' => 'application/pdf',

            'Content-Disposition' =>

                'attachment; filename="' . $filename . '"',

        ]);

    }

}
