<?php

namespace Rizky92\Xlswriter;

use Exception;
use Illuminate\Support\Collection;
use Illuminate\Support\Str;
use Illuminate\Support\Facades\Storage;
use Vtiful\Kernel\Excel;

class ExcelExport
{
    /**
     * The excel instance
     */
    protected Excel $excel;

    /**
     * Array of columns for the data
     * 
     * @var string[]
     */
    protected array $columnHeaders = [];

    /**
     * Array of titles for page header
     * 
     * @var string[]
     */
    protected array $pageHeaders = [];

    /**
     * The exported excel file name
     */
    protected string $filename;

    /**
     * Array of configurations for excel instance
     * 
     * @var string[]
     */
    protected array $config = [];

    /**
     * Array of sheet names for excel
     * 
     * @var string[]
     */
    protected array $sheets = [];

    /**
     * Base path for exported excel file
     */
    protected string $basePath = 'excel';

    /**
     * The disk storage driver used to export the excel
     */
    protected string $disk = 'public';

    /**
     * The stored data instance
     * 
     * @var \Illuminate\Support\Collection<array-key, mixed>|array<array-key, mixed>
     */
    protected $data;

    /**
     * Initialize a new object
     */
    public function __construct(
        string $filename,
        string $sheetName = 'Sheet 1',
        string $basePath = 'excel/',
        string $disk = 'public'
    ) {
        $this->filename = Str::of($filename)
            ->trim('/')
            ->remove('excel/');

        $this->basePath = $basePath;

        $this->disk = $disk;

        $this->sheets[0] = $sheetName;

        $this->configure('path', Storage::disk($disk)->path($basePath));

        $this->excel = (new Excel($this->config))
            ->fileName($this->filename, $this->sheets[0]);
    }

    /**
     * Set the column headers to display for the cell data
     * 
     * @param  string[] $columnHeaders
     */
    public function setColumnHeaders(array $columnHeaders = []): self
    {
        $this->columnHeaders = $columnHeaders;

        return $this;
    }

    /**
     * Set the page headers for the given sheet or all sheets
     * 
     * @param  string[] $pageHeaders
     */
    public function setPageHeaders(array $pageHeaders = []): self
    {
        $this->pageHeaders = $pageHeaders;

        return $this;
    }

    /**
     * Set the type of data to be inserted to excel
     * 
     * @param  \Illuminate\Support\Collection<array-key, mixed>|array<array-key, mixed> $data
     */
    public function setData($data = []): self
    {
        if ($data instanceof Collection) {
            $data = $data->toArray();
        }

        $this->putColumnHeadersToCell();
        $this->putPageHeadersToCell();

        $this->excel->data($data);

        return $this;
    }

    /**
     * Add a new sheet to excel
     */
    public function addSheet(string $sheetName): self
    {
        $this->excel
            ->addSheet($sheetName)
            ->checkoutSheet($sheetName);

        $this->sheets = array_merge($this->sheets, [$sheetName]);

        return $this;
    }

    /**
     * Use currently available sheets
     * 
     * @throws \Exception
     */
    public function useSheet(string $sheetName): self
    {
        if (! in_array($sheetName, $this->sheets, true)) {
            throw new Exception(sprintf(
                'No sheets are available for sheet %s.',
                [$sheetName]
            ));
        }

        $this->excel->checkoutSheet($sheetName);

        return $this;
    }

    /**
     * Get currently available sheets
     * 
     * @return string[]
     */
    public function getAvailableSheets(): array
    {
        return $this->sheets;
    }

    /**
     * Set the disk for excel export
     */
    public function setDisk(string $name = 'public'): self
    {
        $this->disk = $name;

        $this->configure('path', Storage::disk($this->disk)->path($this->basePath));

        return $this;
    }

    /**
     * Set the base path for the excel export
     */
    public function setBasePath(string $basePath = 'excel'): self
    {
        $this->basePath = $basePath;

        $this->configure('path', Storage::disk($this->disk)->path($this->basePath));

        return $this;
    }

    /** 
     * Save excel to disk and return its output path relative to storage
     */
    public function save(): string
    {
        $exportedFilename = $this->basePath . '/' . $this->filename;

        if (!Storage::disk($this->disk)->exists($exportedFilename)) {
            $this->excel->output();
        }

        return $exportedFilename;
    }

    /**
     * Export excel to file as downloadable
     * 
     * @return ?\Symfony\Component\HttpFoundation\StreamedResponse
     */
    public function export()
    {
        $exportedFilename = $this->basePath . '/' . $this->filename;

        if (!Storage::disk($this->disk)->exists($exportedFilename)) {
            $this->excel->output();
        }

        return Storage::disk($this->disk)->download($exportedFilename);
    }

    /**
     * Create a new instance
     * 
     * @param  string $filename
     * @param  string $sheetName
     * @param  string $basePath
     * @param  string $disk
     * 
     * @return static
     */
    public static function make(
        string $filename,
        string $sheetName = 'Sheet 1',
        string $basePath = 'excel/',
        string $disk = 'public'
    ) {
        return new static($filename, $sheetName, $basePath, $disk);
    }

    protected function configure(string $key, $value): void
    {
        $this->config[$key] = $value;
    }

    /**
     * @throws \Exception
     */
    protected function putColumnHeadersToCell(): void
    {
        if (empty($this->columnHeaders)) {
            throw new Exception("Cell column headers need to be set first!");
        }

        if (empty($this->pageHeaders)) {
            $this->excel->header($this->columnHeaders);

            return;
        }

        foreach (array_values($this->columnHeaders) as $id => $column) {
            $this->excel->insertText(count($this->pageHeaders), $id, $column);
        }

        $this->excel->insertText(count($this->pageHeaders) + 1, 0, '');
    }

    protected function putPageHeadersToCell(): void
    {
        if (empty($this->pageHeaders)) {
            return;
        }

        $colStart = $colEnd = 'A';

        for ($i = 0; $i < count($this->columnHeaders) - 1; $i++) {
            $colEnd++;
        }

        foreach (array_values($this->pageHeaders) as $id => $title) {
            $i = $id + 1;

            $this->excel->mergeCells("{$colStart}{$i}:{$colEnd}{$i}", $title);
        }
    }
}
