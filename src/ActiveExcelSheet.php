<?php
namespace codemix\excelexport;

use yii\helpers\ArrayHelper;

/**
 * An excel sheet that is rendered with data from an `ActiveQuery`.
 * A query must be set with `setQuery()`.
 */
class ActiveExcelSheet extends ExcelSheet
{
    /**
     * @var string the format string for `date` columns
     */
    public $dateFormat = 'dd/mm/yyyy';

    /**
     * @var string the format string for `datetime` columns
     */
    public $dateTimeFormat = 'dd/mm/yyyy hh:mm:ss';

    /**
     * @var int the number of records to fetch in batches. Default is `100`.
     */
    public $batchSize = 100;

    protected $_query;
    protected $_attributes;
    protected $_columnTypes;

    /**
     * @return yii\db\ActiveQuery the query for the sheet data
     */
    public function getQuery()
    {
        if ($this->_query===null) {
            throw new \Exception('No query set');
        }
        return $this->_query;
    }

    /**
     * @param yii\db\ActiveQuery $value the query for the sheet data
     */
    public function setQuery($value)
    {
        $this->_query = $value;
    }

    /**
     * @return \yii\db\BatchQueryResult the row records in batches of `$batchSize`
     */
    public function getData()
    {
        return $this->getQuery()->each($this->batchSize);
    }

    /**
     * @return string[] list of attributes for the table columns. If no attributes are set,
     * attributes are set to `ActiveRecord::attributes()` for the main query record.
     */
    public function getAttributes()
    {
        if ($this->_attributes===null) {
            $class = $this->getQuery()->modelClass;
            $model = new $class;
            $this->_attributes = $model->attributes();
        }
        return $this->_attributes;
    }

    /**
     * @param string[] $value list of attributes for the table columns
     */
    public function setAttributes($value)
    {
        $this->_attributes = $value;
    }

    /**
     * @inheritdoc
     */
    public function setData($value)
    {
        throw new \Exception('Data can not be set on ActiveExcelSheet');
    }

    /**
     * @return string[] the column titles. If not set, the respective attribute label is used
     */
    public function getTitles()
    {
        if ($this->_titles===null) {
            $class = $this->getQuery()->modelClass;
            $model = new $class;
            $this->_titles = array_map(function ($a) use ($model) {
                return $model->getAttributeLabel($a);
            }, $this->getAttributes());
        }
        return $this->_titles;
    }

    /**
     * @param string[]|false $value the column titles indexed by 0-based column index.
     * The array is merged with the default titles from `getTitles()` (=attribute labels).
     * If an empty array or `false`, no titles will be generated.
     */
    public function setTitles($value)
    {
        if (!$value) {
            $this->_titles = $value;
        } else {
            if ($this->_titles===null) {
                $this->getTitles(); // Sets attribute labels as defaults
            }
            foreach ($value as $i => $v) {
                $this->_titles[$i] = $v;
            }
        }
    }

    /**
     * @param string[] $value the format strings for the column cells indexed by 0-based column index.
     * If not set, the formats are auto-generated from the DB column types.
     */
    public function getFormats()
    {
        if ($this->_formats===null) {
            $this->_formats = [];
            $attrs = $this->getAttributes();
            $types = $this->getColumnTypes();
            foreach ($attrs as $c => $attr) {
                switch ($types[$c]->type) {
                    case 'date':
                        $this->_formats[$c] = $this->dateFormat;
                        break;
                    case 'datetime':
                        $this->_formats[$c] = $this->dateTimeFormat;
                        break;
                    case 'decimal':
                        $decimals = str_pad('#,', $types[$c]->scale, '#');
                        $zeroPad = str_pad('0.', $types[$c]->scale, '0');
                        $this->_formats[$c] = $decimals.$zeroPad;
                        break;
                }
            }
        }
        return $this->_formats;
    }

    /**
     * @param string[]|false $value the format strings for the column cells indexed by 0-based column index.
     * The array is merged with the default formats from `getFormats()` (auto-generated from DB columns).
     * If an empty array or `false`, no formats are applied.
     */
    public function setFormats($value)
    {
        if (!$value) {
            $this->_formats = $value;
        } else {
            if ($this->_formats===null) {
                $this->getFormats(); // Sets auto-generated formats as defaults
            }
            foreach ($value as $i => $v) {
                $this->_formats[$i] = $v;
            }
        }
    }

    /**
     * @return Callable[] the value formatters for the column cells indexed by 0-based column index.
     * If not set, the formatters are aut-generated from the DB column types.
     */
    public function getFormatters()
    {
        if ($this->_formatters===null) {
            $this->_formatters = [];
            $attrs = $this->getAttributes();
            $types = $this->getColumnTypes();
            foreach ($attrs as $c => $attr) {
                switch ($types[$c]->type) {
                    case 'date':
                    case 'datetime':
                        $this->_formatters[$c] = function ($v) {
                            return \PHPExcel_Shared_Date::PHPToExcel(strtotime($v));
                        };
                        break;
                }
            }
        }
        return $this->_formatters;
    }

    /**
     * @param Callable[]|null $value the value formatters for the column cells indexed by 0-based column index.
     * The array is merged with the default formats from `getFormatters()` (auto-generated from DB columns).
     * If an empty array or `false`, no formatters are applied.
     */
    public function setFormatters($value)
    {
        if (!$value) {
            $this->_formatters = $value;
        } else {
            if ($this->_formatters===null) {
                $this->getFormatters(); // Sets auto-generated formatters as defaults
            }
            foreach ($value as $i => $v) {
                $this->_formatters[$i] = $v;
            }
        }
    }

    /**
     * @return yii\db\ColumnSchema[] the DB column types `ColumnSchema::$type` indexed by 0-based column index
     */
    protected function getColumnTypes()
    {
        if ($this->_columnTypes===null) {
            $class = $this->getQuery()->modelClass;
            $model = new $class;
            $this->_columnTypes = array_map(function ($attr) use ($model) {
                return self::getType($model, $attr);
            }, $this->getAttributes());

        }
        return $this->_columnTypes;
    }

    /**
     * @inheritdoc
     */
    protected function renderRow($data, $row, $formats, $formatters, $callbacks, $types)
    {
        $values = array_map(function ($attr) use ($data) {
            return ArrayHelper::getValue($data, $attr);
        }, $this->getAttributes());
        return parent::renderRow($values, $row, $formats, $formatters, $callbacks, $types);
    }

    /**
     * Returns either the ColumnSchema or a new instance of the related model for the given attribute name.
     *
     * The name can be specified in dot format, like `company.name` in which case the ColumnSchema for
     * the `name` attribute in the related `company` record would be returned.
     *
     * If the attribute is a relation name (which could also use dot notation) then `$isRelation` must
     * be set to `true`. In this case an instance of the related ActiveRecord class is returned.
     *
     * @param yii\db\ActiveRecord $model the model where the attribute exist
     * @param string $attribute name of the attribute
     * @param mixed $isRelation whether the name specifies a relation, in which case an `ActiveRecord`
     * is returned. Default is `false`, which returns a `ColumnSchema`.
     * @return yii\db\ColumnSchema|yii\db\ActiveRecord the type instance of the attribute
     */
    public static function getType($model, $attribute, $isRelation = false)
    {
        if (($pos = strrpos($attribute, '.')) !== false) {
            $model = self::getType($model, substr($attribute, 0, $pos), true);
            $attribute = substr($attribute, $pos + 1);
        }
        if ($isRelation) {
            $class = $model->getRelation($attribute)->modelClass;
            return new $class;
        } else {
            return $model->getTableSchema()->columns[$attribute];
        }
    }
}
