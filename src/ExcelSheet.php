<?php
namespace codemix\excelexport;

use yii\base\Object;

/**
 * An excel worksheet
 */
class ExcelSheet extends Object
{
    protected $_sheet;
    protected $_data;
    protected $_titles;
    protected $_types;
    protected $_formats;
    protected $_formatters;
    protected $_callbacks;
    protected $_row = 1;

    /**
     * @param PHPExcel_WorkSheet $sheet
     * @param array $config
     */
    public function __construct($sheet, $config = [])
    {
        parent::__construct($config);
        $this->_sheet = $sheet;
    }

    /**
     * @return PHPExcel_WorkSheet
     */
    public function getSheet()
    {
        return $this->_sheet;
    }

    /**
     * @return array|\Iterator the data for the rows of the sheet
     */
    public function getData()
    {
        return $this->_data;
    }

    /**
     * @param array|\Iterator $value the data for the rows of the sheet
     */
    public function setData($value)
    {
        $this->_data = $value;
    }

    /**
     * @return string[]|null|false the column titles indexed by 0-based column index.
     * If empty, `null` or `false`, no titles will be generated.
     */
    public function getTitles()
    {
        return $this->_titles;
    }

    /**
     * @param string[]|null|false $value the column titles indexed by 0-based column index.
     * If empty or `false`, no titles will be generated.
     */
    public function setTitles($value)
    {
        $this->_titles = $value;
    }

    /**
     * @return string[]|null the types for the column cells indexed by 0-based column index. See the
     * `PHPExcel_Cell_DataType::TYPE_*` constants for available types. If no type is set for a column,
     * PHPExcel will autodetect the correct type.
     */
    public function getTypes()
    {
        return $this->_types;
    }

    /**
     * @param string[]|null $value the types for the column cells indexed by 0-based column index
     */
    public function setTypes($value)
    {
        $this->_types = $value;
    }

    /**
     * @return string[]|null the format strings for the column cells indexed by 0-based column index
     */
    public function getFormats()
    {
        return $this->_formats;
    }

    /**
     * @param string[]|null $value the format strings for the column cells indexed by 0-based column index
     */
    public function setFormats($value)
    {
        $this->_formats = $value;
    }

    /**
     * @return Callable[]|null the value formatters for the column cells indexed by 0-based column index.
     * The function signature is `function ($value, $row, $data)` where `$value` is the cell value, `$row`
     * is the row index and `$data` is the row data.
     */
    public function getFormatters()
    {
        return $this->_formatters;
    }

    /**
     * @param Callable[]|null $value the value formatters for the column cells indexed by 0-based column index
     */
    public function setFormatters($value)
    {
        $this->_formatters = $value;
    }

    /**
     * @return Callable[]|null column callbacks indexed by 0-based column index that get called after rendering a cell.
     * The function signature is `function ($cell, $column, $row)` where `$cell` is the `PHPExcel_Cell` object and
     * `$row` and `$column` are the row and column index.
     */
    public function getCallbacks()
    {
        return $this->_callbacks;
    }

    /**
     * @param Callable[]|null $value callbacks that get called after rendering a column cell indexed by 0-based column index.
     */
    public function setCallbacks($value)
    {
        $this->_callbacks = $value;
    }

    /**
     * Render the sheet
     */
    public function render()
    {
        $this->renderTitle();
        $this->renderRows();
    }

    /**
     * Render the title row if any
     */
    protected function renderTitle()
    {
        $titles = $this->getTitles();
        if ($titles) {
            $col = 0;
            foreach ($titles as $title) {
                $this->_sheet->setCellValueByColumnAndRow($col++, $this->_row, $title);
            }
            $this->_row++;
        }
    }

    /**
     * Render the data rows if any
     */
    protected function renderRows()
    {
        $formats = self::normalizeIndex($this->getFormats());
        $formatters = self::normalizeIndex($this->getFormatters());
        $callbacks = self::normalizeIndex($this->getCallbacks());
        $types = self::normalizeIndex($this->getTypes());

        foreach ($this->getData() as $data) {
            $this->renderRow($data, $this->_row++, $formats, $formatters, $callbacks, $types);
        }
    }

    /**
     * Render a single row
     *
     * @param array $data the row data
     * @param int $row the index of the current row
     * @param mixed $formats formats with normalized index
     * @param mixed $formatters formatters with normalized index
     * @param mixed $callbacks callbacks with normalized index
     * @param mixed $types types with normalized index
     */
    protected function renderRow($data, $row, $formats, $formatters, $callbacks, $types)
    {
        foreach (array_values($data) as $i => $value) {
            if (isset($formatters[$i]) && is_callable($formatters[$i])) {
                $value = call_user_func($formatters[$i], $value, $row, $data);
            }
            if (isset($types[$i])) {
                $this->_sheet->setCellValueExplicitByColumnAndRow($i, $row, $value, $types[$i]);
            } else {
                $this->_sheet->setCellValueByColumnAndRow($i, $row, $value);
            }
            if (isset($formats[$i])) {
                $this->_sheet->getStyleByColumnAndRow($i, $row)
                    ->getNumberFormat()
                    ->setFormatCode($formats[$i]);
            }
            if (isset($callbacks[$i]) && is_callable($callbacks[$i])) {
                call_user_func($callbacks[$i], $this->_sheet->getCellByColumnAndRow($i, $row), $i, $row);
            }
        }
    }

    /**
     * @param array $data
     * @return array the array with alphanumeric column keys (A, B, C, ...) converted to numeric indices
     */
    protected static function normalizeIndex($data)
    {
        if (!is_array($data)) {
            return $data;
        }
        $result = [];
        foreach ($data as $k => $v) {
            $result[is_string($k) ? \PHPExcel_Cell::columnIndexFromString($k)-1 : $k] = $v;
        }
        return $result;
    }
}
