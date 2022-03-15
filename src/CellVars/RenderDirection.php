<?php

namespace Kaxiluo\PhpExcelTemplate\CellVars;

class RenderDirection
{
    protected static $cache = [];

    const RIGHT = 'right';
    const DOWN = 'down';

    public static function isValid($value): bool
    {
        return in_array($value, static::toArray(), true);
    }

    public static function toArray()
    {
        $class = \get_called_class();
        if (!isset(static::$cache[$class])) {
            $reflection = new \ReflectionClass($class);
            static::$cache[$class] = $reflection->getConstants();
        }

        return static::$cache[$class];
    }
}
