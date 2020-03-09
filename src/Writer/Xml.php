<?php

namespace Eggpan\PHPXlsxWriter\Writer;

use XMLWriter;

class Xml
{
    protected static $writer;
    protected static $instance;

    protected static function init()
    {
        static::$instance = new static();
    }

    public function __construct()
    {
        static::$writer = new XMLWriter();
        static::$writer->openMemory();
    }

    public static function element($element, $attributes = null, $text = null)
    {
        empty(static::$instance) and static::init();

        return static::start($element, $attributes, $text)
            ->end();
    }

    public static function end()
    {
        empty(static::$instance) and static::init();

        static::$writer->endElement();
        return static::$instance;
    }

    public static function getWriter()
    {
        empty(static::$instance) and static::init();

        return static::$writer;
    }

    public static function start($element , $attributes = null, $text = null)
    {
        empty(static::$writer) and static::init();

        static::$writer->startElement($element);
        if (isset($attributes)) {
            foreach ($attributes as $name => $value) {
                static::$writer->writeAttribute($name, $value);
            }
        }
        if (isset($text)) {
            static::$writer->text($text);
        }
        return static::$instance;
    }
}
