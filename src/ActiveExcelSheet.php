<?php
namespace codemix\excelexport;

use Yii;
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
    protected $_columnSchemas;
    protected $_modelInstance;

    /**
     * @return yii\db\ActiveQuery the query for the sheet data
     */
    public function getQuery()
    {
        if ($this->_query === null) {
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
     * @return string[] list of attributes for the table columns. If no
     * attributes are set, attributes are set to `ActiveRecord::attributes()`
     * for the main query record.
     */
    public function getAttributes()
    {
        if ($this->_attributes === null) {
            $this->_attributes = $this->getModelInstance()->attributes();
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
     * @return string[] the column titles. If not set, the respective attribute
     * label is used
     */
    public function getTitles()
    {
        if ($this->_titles === null) {
            $model = $this->getModelInstance();
            $this->_titles = array_map(function ($a) use ($model) {
                return $model->getAttributeLabel($a);
            }, $this->getAttributes());
        }
        return $this->_titles;
    }

    /**
     * @param string[]|false $value the column titles indexed by 0-based column
     * index.  The array is merged with the default titles from `getTitles()`
     * (=attribute labels).  If an empty array or `false`, no titles will be
     * generated.
     */
    public function setTitles($value)
    {
        if (!$value) {
            $this->_titles = $value;
        } else {
            if ($this->_titles === null) {
                $this->getTitles(); // Sets attribute labels as defaults
            }
            foreach ($value as $i => $v) {
                $this->_titles[$i] = $v;
            }
        }
    }

    /**
     * @param string[] $value the format strings for the column cells indexed
     * by 0-based column index.  If not set, the formats are auto-generated
     * from the DB column types.
     */
    public function getFormats()
    {
        if ($this->_formats === null) {
            $this->_formats = [];
            $attrs = $this->normalizeIndex($this->getAttributes());
            $schemas = $this->normalizeIndex($this->getColumnSchemas());
            foreach ($attrs as $c => $attr) {
                if (!isset($schemas[$c])) {
                    continue;
                }
                switch ($schemas[$c]->type) {
                    case 'date':
                        $this->_formats[$c] = $this->dateFormat;
                        break;
                    case 'datetime':
                        $this->_formats[$c] = $this->dateTimeFormat;
                        break;
                    case 'decimal':
                        $decimals = str_pad('#,', $schemas[$c]->scale, '#');
                        $zeroPad = str_pad('0.', $schemas[$c]->scale, '0');
                        $this->_formats[$c] = $decimals.$zeroPad;
                        break;
                }
            }
        }
        return $this->_formats;
    }

    /**
     * @param string[]|false $value the format strings for the column cells
     * indexed by 0-based column index.  The array is merged with the default
     * formats from `getFormats()` (auto-generated from DB columns).  If an
     * empty array or `false`, no formats are applied.
     */
    public function setFormats($value)
    {
        if (!$value) {
            $this->_formats = $value;
        } else {
            if ($this->_formats === null) {
                $this->getFormats(); // Sets auto-generated formats as defaults
            }
            foreach ($value as $i => $v) {
                $this->_formats[$i] = $v;
            }
        }
    }

    /**
     * @return Callable[] the value formatters for the column cells indexed by
     * 0-based column index.  If not set, the formatters are aut-generated from
     * the DB column types.
     */
    public function getFormatters()
    {
        if ($this->_formatters === null) {
            $this->_formatters = [];
            $attrs = $this->normalizeIndex($this->getAttributes());
            $schemas = $this->normalizeIndex($this->getColumnSchemas());
            foreach ($attrs as $c => $attr) {
                if (!isset($schemas[$c])) {
                    continue;
                }
                switch ($schemas[$c]->type) {
                    case 'date':
                        $this->_formatters[$c] = function ($v) {
                            if (empty($v)) {
                                return null;
                            } else {
                                // Set the correct timezone before converting to a UNIX timestamp.
                                // This prevents dates from being altered due to timezone
                                // conversion, e.g.
                                // '2017-12-05 00:00:00' could become
                                // '2017-12-04 23:00:00'
                                $timezone = date_default_timezone_get();
                                date_default_timezone_set(Yii::$app->formatter->defaultTimeZone);
                                $timestamp = strtotime($v);
                                date_default_timezone_set($timezone);
                                return \PHPExcel_Shared_Date::PHPToExcel($timestamp);
                            }
                        };
                        break;
                    case 'datetime':
                        $this->_formatters[$c] = function ($v) {
                            if (empty($v)) {
                                return null;
                            } else {
                                return \PHPExcel_Shared_Date::PHPToExcel($this->toExcelTime($v));
                            }
                        };
                        break;
                }
            }
        }
        return $this->_formatters;
    }

    /**
     * @param Callable[]|null $value the value formatters for the column cells
     * indexed by 0-based column index.  The array is merged with the default
     * formats from `getFormatters()` (auto-generated from DB columns).  If an
     * empty array or `false`, no formatters are applied.
     */
    public function setFormatters($value)
    {
        if (!$value) {
            $this->_formatters = $value;
        } else {
            if ($this->_formatters === null) {
                $this->getFormatters(); // Sets auto-generated formatters as defaults
            }
            foreach ($value as $i => $v) {
                $this->_formatters[$i] = $v;
            }
        }
    }

    /**
     * @return yii\db\ActiveRecord an instance of the main model on which the
     * query is performed on. This is used to obtain column titles and types.
     */
    public function getModelInstance()
    {
        if ($this->_modelInstance === null) {
            $class = $this->getQuery()->modelClass;
            $this->_modelInstance = new $class;
        }
        return $this->_modelInstance;
    }

    /**
     * @param yii\db\ActiveRecord $model an instance of the main model on which
     * the query is performed on. This is used to obtain column titles and
     * types.
     */
    public function setModelInstance($model)
    {
        $this->_modelInstance = $model;
    }

    /**
     * @return yii\db\ActiveRecord a new instance of a related model for the
     * given model. This is used to obtain column types.
     */
    protected static function getRelatedModelInstance($model, $name)
    {
        $class = $model->getRelation($name)->modelClass;
        return new $class;
    }

    /**
     * @return yii\db\ColumnSchema[] the DB column schemas indexed by 0-based
     * column index. This only includes columns for which a DB schema exists.
     */
    protected function getColumnSchemas()
    {
        if ($this->_columnSchemas === null) {
            $model = $this->getModelInstance();
            $schemas = array_map(function ($attr) use ($model) {
                return self::getSchema($model, $attr);
            }, $this->getAttributes());
            // Filter out null values
            $this->_columnSchemas = array_filter($schemas);
        }
        return $this->_columnSchemas;
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
     * Convert a datetime to the right excel timestamp
     *
     * This method will use [[\yii\i18n\Formatter::defaultTimeZone]] and
     * [[\yii\base\Application::timeZone]] to convert the given datetime
     * from DB to application timezone.
     *
     * @param string $value the datetime value
     * @return int timezone offset in seconds 
     * @see [[yii\i18n\Formatter::defaultTimezone]]
     * @see [[yii\i18n\Formatter::timezone]]
     */
    protected function toExcelTime($value)
    {
        // "Cached" timezone instances
        static $defaultTimezone;
        static $timezone;

        if (Yii::$app->formatter->defaultTimeZone === Yii::$app->timeZone) {
            return strtotime($value);
        } else {
            if ($timezone === null) {
                $defaultTimezone = new \DateTimeZone(Yii::$app->formatter->defaultTimeZone);
                $timezone = new \DateTimeZone(Yii::$app->timeZone);
            }

            // Offset can depend on given datetime due to DST
            $defaultDatetime = new \DateTime($value, $defaultTimezone);
            $offset = $timezone->getOffset($defaultDatetime);

            // PHPExcel_Shared_Date::PHPToExcel() method expects a
            // "pseudo-timestamp": Something like a UNIX timestamp but
            // including local timezone offset.
            return $defaultDatetime->getTimestamp() + $offset;
        }
    }

    /**
     * Returns either the ColumnSchema or a new instance of the related model
     * for the given attribute name.
     *
     * The name can be specified in dot format, like `company.name` in which
     * case the ColumnSchema for the `name` attribute in the related `company`
     * record would be returned.
     *
     * If the attribute is a relation name (which could also use dot notation)
     * then `$isRelation` must be set to `true`. In this case an instance of
     * the related ActiveRecord class is returned.
     *
     * @param yii\db\ActiveRecord $model the model where the attribute exist
     * @param string $attribute name of the attribute
     * @param mixed $isRelation whether the name specifies a relation, in which
     * case an `ActiveRecord` is returned. Default is `false`, which returns a
     * `ColumnSchema`.
     * @return yii\db\ColumnSchema|yii\db\ActiveRecord|null the type instance
     * of the attribute or `null` if the attribute is not a DB column (e.g.
     * public property or defined by getter)
     */
    public static function getSchema($model, $attribute, $isRelation = false)
    {
        if (($pos = strrpos($attribute, '.')) !== false) {
            $model = self::getSchema($model, substr($attribute, 0, $pos), true);
            $attribute = substr($attribute, $pos + 1);
        }
        if ($isRelation) {
            return self::getRelatedModelInstance($model, $attribute);
        } else {
            $columnSchemas = $model->getTableSchema()->columns;
            return isset($columnSchemas[$attribute]) ? $columnSchemas[$attribute] : null;
        }
    }
}
