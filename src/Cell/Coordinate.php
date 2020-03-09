<?php

namespace Eggpan\PHPXlsxWriter\Cell;

use Eggpan\PHPXlsxWriter\Exceptions\UnknownErrorException;

class Coordinate
{
    /**
     * カラム数値からカラム文字列を取得する
     *
     * @param int $index
     * @return string
     */
    public static function stringFromColumnIndex(int $index): string
    {
        static $cache = [];

        if (empty($cache[$index])) {
            $name1 = $index <= 702 ? '' : chr((floor(($index - 27) / 676) - 1) % 26 + 65);
            $name2 = $index <= 26 ? '' : chr((floor(($index - 1) / 26) - 1) % 26 + 65);
            $name3 = chr((($index - 1) % 26) + 65);
            $cache[$index] = $name1.$name2.$name3;
        }
        return $cache[$index];
    }

    /**
     * カラム文字列からカラム数値を取得する
     *
     * @param string $column
     * @return int
     */
    public static function columnIndexFromString(string $column): int
    {
        static $cache = [];

        $column = strtoupper($column);
        if (empty($cache[$column])) {
            switch (strlen($column)) {
                case 1:
                    $cache[$column] = ord($column) - 64;
                break;
                case 2:
                    $cache[$column] = (ord($column[0]) - 64) * 26 + (ord($column[1]) - 64);
                break;
                case 3:
                    $cache[$column] = (ord($column[0]) - 64) * 676 + (ord($column[1]) - 64) * 26 + (ord($column[2]) - 64);
                break;
            }
        }

        return $cache[$column];
    }

    // ----------------------------------------------------------------

    /**
     * 文字列のセルを元にカラム番号と行番号を返す
     *
     * @param string $cell
     * @return array
     */
    public static function columnRowIndexFromString(string $cell): array
    {
        if (preg_match('/([A-Z]+)([0-9]+)/i', $cell, $match) !== 1) {
            throw new UnknownErrorException();
        }
        $column = Coordinate::columnIndexFromString($match[1]);
        $row = (int)$match[2];
        return [$column, $row];
    }
}
