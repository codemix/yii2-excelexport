<?php
namespace codemix\excelexport;

use Yii;
use yii\base\Object;
use mikehaertl\tmp\File;

/**
 * This class represents an excel file.
 */
class ExcelFile extends Object
{
    /**
     * @var string the writer class to use. Default is `\PHPExcel_Writer_Excel2007`.
     */
    public $writer = '\PHPExcel_Writer_Excel2007';

    protected $_workbook;
    protected $_sheets;
    protected $_tmpFile;
    protected $_created = false;

    /**
     * @return PHPExcel the workbook instance
     */
    public function getWorkbook()
    {
        if ($this->_workbook===null) {
            $this->_workbook = new \PHPExcel();
        }
        return $this->_workbook;
    }

    /**
     * @return mikehaertl\tmp\File the instance of the temporary excel file
     */
    public function getTmpFile()
    {
        if ($this->_tmpFile===null) {
            $this->_tmpFile = new File('');
        }
        return $this->_tmpFile;
    }

    /**
     * @return array the sheet configuration
     */
    public function getSheets()
    {
        return $this->_sheets;
    }

    /**
     * @param array $value the sheet configuration. This must be an array where keys
     * are sheet names and values are arrays with the configuration options for an
     * instance if `ExcelSheet`.
     */
    public function setSheets($value)
    {
        $this->_sheets = $value;
    }

    /**
     * Save the file under the given name
     *
     * @param string $filename
     * @return bool whether the file was saved successfully
     */
    public function saveAs($filename)
    {
        $this->createFile();
        return $this->getTmpFile()->saveAs($filename);
    }

    /**
     * Send the Excel file for download
     *
     * @param string|null $filename the filename to send. If empty, the file is streamed inline.
     * @param bool $inline whether to force inline display of the file, even if filename is present.
     */
    public function send($filename = null, $inline = false)
    {
        $this->createFile();
        $this->getTmpFile()->send($filename, 'application/vnd.ms-excel', $inline);
    }

    /**
     * Create the Excel file and save it to the temp file
     */
    protected function createFile()
    {
        if (!$this->_created) {
            $workbook = $this->getWorkbook();
            $i = 0;
            foreach ($this->sheets as $title => $config) {
                if (is_string($config)) {
                    $config = ['class' => $config];
                } elseif (is_array($config)) {
                    if (!isset($config['class'])) {
                        $config['class'] = ExcelSheet::className();
                    }
                } elseif (!is_object($config)) {
                    throw new \Exception('Invalid sheet configuration');
                }
                $sheet = (0===$i++) ? $workbook->getActiveSheet() : $workbook->createSheet();
                if (is_string($title)) {
                    $sheet->setTitle($title);
                }
                Yii::createObject($config, [$sheet])->render();
            }
            $class = $this->writer;
            $writer = new $class($workbook);
            $writer->save((string) $this->getTmpFile());
            $this->_created = true;
        }
    }
}
