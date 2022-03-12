<?php

namespace Kaxiluo\PhpExcelTemplate\CellVars;

class RenderDirection
{
    protected static $cache = [];

    private $direction;

    const RIGHT = 'right';
    const DOWN = 'down';

    public function __construct(string $direction)
    {
        if (!static::isValid($direction)) {
            throw new \UnexpectedValueException('Unexpected RenderDirection [' . $direction . ']');
        }
        $this->direction = $direction;
    }

    public function getDirection(): string
    {
        return $this->direction;
    }

    public function isDirection($value): bool
    {
        return $this->direction === $value;
    }

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

    public function __toString()
    {
        return $this->direction;
    }
}
