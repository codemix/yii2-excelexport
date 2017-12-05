<?php
namespace codemix\excelexport;

use yii\base\Component;

/**
 * An excel worksheet
 */
class ExcelSheet extends Component
{
    const EVENT_BEFORE_RENDER = 'beforeRender';
    const EVENT_AFTER_RENDER = 'afterRender';

    /**
     * @var int|string the start column name or its 0-based index. When this is
     * set, the 0-based offset is added to all numeric keys used anywhere in
     * this class. Columns referenced by name will stay unchanged.  Default is
     * 'A'.
     */
    public $startColumn = 'A';

    /**
     * @var int the start row. Default is 1.
     */
    public $startRow = 1;

    protected $_sheet;
    protected $_data;
    protected $_titles;
    protected $_types;
    protected $_formats;
    protected $_formatters;
    protected $_styles = [];
    protected $_callbacks;
    protected $_row;

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
     * @return string[]|null|false the column titles indexed by column name or
     * 0-based index. If empty, `null` or `false`, no titles will be generated.
     */
    public function getTitles()
    {
        return $this->_titles;
    }

    /**
     * @param string[]|null|false $value the column titles indexed by 0-based
     * column index.  If empty or `false`, no titles will be generated.
     */
    public function setTitles($value)
    {
        $this->_titles = $value;
    }

    /**
     * @return string[]|null the types for the column cells indexed by 0-based
     * column index. See the `PHPExcel_Cell_DataType::TYPE_*` constants for
     * available types. If no type is set for a column, PHPExcel will
     * autodetect the correct type.
     */
    public function getTypes()
    {
        return $this->_types;
    }

    /**
     * @param string[]|null $value the types for the column cells indexed by
     * 0-based column index
     */
    public function setTypes($value)
    {
        $this->_types = $value;
    }

    /**
     * @return string[]|null the format strings for the column cells indexed by
     * 0-based column index
     */
    public function getFormats()
    {
        return $this->_formats;
    }

    /**
     * @param string[]|null $value the format strings for the column cells
     * indexed by 0-based column index
     */
    public function setFormats($value)
    {
        $this->_formats = $value;
    }

    /**
     * @return Callable[]|null the value formatters for the column cells
     * indexed by 0-based column index.  The function signature is `function
     * ($value, $row, $data)` where `$value` is the cell value, `$row` is the
     * row index and `$data` is the row data.
     */
    public function getFormatters()
    {
        return $this->_formatters;
    }

    /**
     * @param Callable[]|null $value the value formatters for the column cells
     * indexed by 0-based column index
     */
    public function setFormatters($value)
    {
        $this->_formatters = $value;
    }

    /**
     * @return array style configuration arrays indexed by cell coordinate or
     * cell range, e.g. `A1:Z1000`.
     */
    public function getStyles()
    {
        return $this->_styles;
    }

    /**
     * @param array $value style configuration arrays indexed by cell
     * coordinate or cell range, e.g. `A1:Z1000`.
     */
    public function setStyles($value)
    {
        $this->_styles = $value;
    }

    /**
     * @return Callable[]|null column callbacks indexed by 0-based column index
     * that get called after rendering a cell.  The function signature is
     * `function ($cell, $column, $row)` where `$cell` is the `PHPExcel_Cell`
     * object and `$row` and `$column` are the row and column index.
     */
    public function getCallbacks()
    {
        return $this->_callbacks;
    }

    /**
     * @param Callable[]|null $value callbacks that get called after rendering
     * a column cell indexed by 0-based column index.
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
        $this->beforeRender();
        $this->_row = $this->startRow;
        $this->renderStyles();
        $this->renderTitle();
        $this->renderRows();
        $this->trigger(self::EVENT_AFTER_RENDER);
    }

    /**
     * Trigger the [[EVENT_BEFORE_RENDER]] event
     */
    public function beforeRender()
    {
        $this->trigger(self::EVENT_BEFORE_RENDER);
    }

    /**
     * Trigger the [[EVENT_AFTER_RENDER]] event
     */
    public function afterRender()
    {
        $this->trigger(self::EVENT_AFTER_RENDER);
    }

    /**
     * Render styles
     */
    protected function renderStyles()
    {
        foreach ($this->getStyles() as $i => $style) {
            $this->_sheet->getStyle($i)->applyFromArray($style);
        }
    }

    /**
     * Render the title row if any
     */
    protected function renderTitle()
    {
        $titles = $this->normalizeIndex($this->getTitles());
        if ($titles) {
            $keys = array_keys($titles);
            $col = array_shift($keys);
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
        $formats = $this->normalizeIndex($this->getFormats());
        $formatters = $this->normalizeIndex($this->getFormatters());
        $callbacks = $this->normalizeIndex($this->getCallbacks());
        $types = $this->normalizeIndex($this->getTypes());

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
            $col = $i + self::normalizeColumn($this->startColumn);
            if (isset($formatters[$col]) && is_callable($formatters[$col])) {
                $value = call_user_func($formatters[$col], $value, $row, $data);
            }
            if (isset($types[$col])) {
                $this->_sheet->setCellValueExplicitByColumnAndRow($col, $row, $value, $types[$col]);
            } else {
                $this->_sheet->setCellValueByColumnAndRow($col, $row, $value);
            }
            if (isset($formats[$col])) {
                $this->_sheet
                    ->getStyleByColumnAndRow($col, $row)
                    ->getNumberFormat()
                    ->setFormatCode($formats[$col]);
            }
            if (isset($callbacks[$col]) && is_callable($callbacks[$col])) {
                $cell = $this->_sheet->getCellByColumnAndRow($col, $row);
                call_user_func($callbacks[$col], $cell, $col, $row);
            }
        }
    }

    /**
     * @param array $data any data indexed by 0-based colum index or by column name.
     * @return array the array with alphanumeric column keys (A, B, C, ...)
     * converted to numeric indices
     */
    protected function normalizeIndex($data)
    {
        if (!is_array($data)) {
            return $data;
        }
        $result = [];
        foreach ($data as $k => $v) {
            $result[self::normalizeColumn($k)] = $v;
        }
        return $result;
    }

    /**
     * @param int|string $column the column either as int or as string. If
     * numeric, the startColumn offset will be added.
     * @return int the normalized numeric column index (0-based).
     */
    public function normalizeColumn($column)
    {
        if (is_string($column)) {
            return \PHPExcel_Cell::columnIndexFromString($column) - 1;
        } else {
            return $column + self::normalizeColumn($this->startColumn);
        }
    }
}
