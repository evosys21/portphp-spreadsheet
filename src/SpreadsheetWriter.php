<?php
/**
 * @noinspection PhpUnhandledExceptionInspection
 */

namespace Port\Spreadsheet;

use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\IOFactory;
use Port\Writer;

/**
 * Writes to an Excel file
 *
 * @author David de Boer <david@ddeboer.nl>
 */
class SpreadsheetWriter implements Writer
{
    /**
     * @var string
     */
    protected $filename;

    /**
     * @var null|string
     */
    protected $sheet;

    /**
     * @var string
     */
    protected $type;

    /**
     * @var boolean
     */
    protected $prependHeaderRow;

    /**
     * @var Spreadsheet
     */
    protected $spreadsheet;

    /**
     * @var integer
     */
    protected $row = 1;

    /**
     * @param string $filename File
     * @param string $sheet Sheet title (optional)
     * @param string $type Excel file type (defaults to Xlsx)
     * @param boolean $prependHeaderRow
     */
    public function __construct(string $filename, $sheet = null, $type = 'Xlsx', $prependHeaderRow = false)
    {
        $this->filename = $filename;
        $this->sheet = $sheet;
        $this->type = $type;
        $this->prependHeaderRow = $prependHeaderRow;
    }

    /**
     * {@inheritdoc}
     */
    public function prepare()
    {
        $reader = IOFactory::createReader($this->type);
        if (file_exists($this->filename) && $reader->canRead($this->filename)) {
            $this->spreadsheet = $reader->load($this->filename);
        } else {
            $this->spreadsheet = new Spreadsheet();
        }

        if (null !== $this->sheet) {
            if (!$this->spreadsheet->sheetNameExists($this->sheet)) {
                $this->spreadsheet->createSheet()->setTitle($this->sheet);
            }
            $this->spreadsheet->setActiveSheetIndexByName($this->sheet);
        }
    }

    /**
     * {@inheritdoc}
     */
    public function writeItem(array $item)
    {

        print_r($item);

        $count = count($item);

        if ($this->prependHeaderRow && 1 == $this->row) {
            $headers = array_keys($item);

            for ($i = 0; $i < $count; $i++) {
                $this->spreadsheet->getActiveSheet()->setCellValueByColumnAndRow($i + 1, $this->row, $headers[ $i ]);
            }
            $this->row++;
        }

        $values = array_values($item);

        for ($i = 0; $i < $count; $i++) {
            $this->spreadsheet->getActiveSheet()->setCellValueByColumnAndRow($i + 1, $this->row, $values[ $i ]);
        }

        $this->row++;
    }

    /**
     * {@inheritdoc}
     */
    public function finish()
    {
        $writer = IOFactory::createWriter($this->spreadsheet, $this->type);
        $writer->save($this->filename);
    }
}
