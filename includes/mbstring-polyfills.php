<?php

if (!defined('MB_CASE_UPPER')) {
    define('MB_CASE_UPPER', 0);
}

if (!defined('MB_CASE_LOWER')) {
    define('MB_CASE_LOWER', 1);
}

if (!defined('MB_CASE_TITLE')) {
    define('MB_CASE_TITLE', 2);
}

function isop_exporter_mb_encoding($encoding)
{
    return $encoding ?: 'UTF-8';
}

function isop_exporter_mb_chars($string)
{
    $chars = preg_split('//u', (string) $string, -1, PREG_SPLIT_NO_EMPTY);
    return is_array($chars) ? $chars : str_split((string) $string);
}

if (!function_exists('mb_strlen')) {
    function mb_strlen($string, $encoding = null)
    {
        $encoding = isop_exporter_mb_encoding($encoding);
        if (function_exists('iconv_strlen')) {
            $length = @iconv_strlen($string, $encoding);
            if ($length !== false) {
                return $length;
            }
        }

        return count(isop_exporter_mb_chars($string));
    }
}

if (!function_exists('mb_substr')) {
    function mb_substr($string, $start, $length = null, $encoding = null)
    {
        $encoding = isop_exporter_mb_encoding($encoding);
        if (function_exists('iconv_substr')) {
            $substr = @iconv_substr($string, $start, $length, $encoding);
            if ($substr !== false) {
                return $substr;
            }
        }

        $chars = isop_exporter_mb_chars($string);
        $slice = array_slice($chars, $start, $length);
        return implode('', $slice);
    }
}

if (!function_exists('mb_strtolower')) {
    function mb_strtolower($string, $encoding = null)
    {
        return strtolower((string) $string);
    }
}

if (!function_exists('mb_strtoupper')) {
    function mb_strtoupper($string, $encoding = null)
    {
        return strtoupper((string) $string);
    }
}

if (!function_exists('mb_strpos')) {
    function mb_strpos($haystack, $needle, $offset = 0, $encoding = null)
    {
        $encoding = isop_exporter_mb_encoding($encoding);
        if (function_exists('iconv_strpos')) {
            $position = @iconv_strpos($haystack, $needle, $offset, $encoding);
            if ($position !== false) {
                return $position;
            }
        }

        $haystackChars = isop_exporter_mb_chars($haystack);
        $needleChars = isop_exporter_mb_chars($needle);
        $needleLength = count($needleChars);

        for ($i = $offset; $i <= count($haystackChars) - $needleLength; $i++) {
            if (array_slice($haystackChars, $i, $needleLength) === $needleChars) {
                return $i;
            }
        }

        return false;
    }
}

if (!function_exists('mb_stripos')) {
    function mb_stripos($haystack, $needle, $offset = 0, $encoding = null)
    {
        return mb_strpos(mb_strtolower($haystack, $encoding), mb_strtolower($needle, $encoding), $offset, $encoding);
    }
}

if (!function_exists('mb_convert_encoding')) {
    function mb_convert_encoding($string, $to_encoding, $from_encoding = null)
    {
        $from_encoding = $from_encoding ?: 'UTF-8';
        if (is_array($from_encoding)) {
            $from_encoding = implode(',', $from_encoding);
        }

        if (function_exists('iconv')) {
            $converted = @iconv($from_encoding, $to_encoding . '//IGNORE', $string);
            if ($converted !== false) {
                return $converted;
            }
        }

        return (string) $string;
    }
}

if (!function_exists('mb_convert_case')) {
    function mb_convert_case($string, $mode, $encoding = null)
    {
        switch ($mode) {
            case MB_CASE_UPPER:
                return mb_strtoupper($string, $encoding);
            case MB_CASE_LOWER:
                return mb_strtolower($string, $encoding);
            case MB_CASE_TITLE:
                return ucwords(mb_strtolower($string, $encoding));
            default:
                return (string) $string;
        }
    }
}

if (!function_exists('mb_check_encoding')) {
    function mb_check_encoding($value = null, $encoding = null)
    {
        if ($value === null) {
            return true;
        }

        if (is_array($value)) {
            foreach ($value as $item) {
                if (!mb_check_encoding($item, $encoding)) {
                    return false;
                }
            }

            return true;
        }

        $string = (string) $value;
        $encoding = strtoupper(isop_exporter_mb_encoding($encoding));

        if ($encoding === 'ASCII') {
            return preg_match('/^[\x00-\x7F]*$/', $string) === 1;
        }

        if ($encoding === 'UTF-8') {
            return preg_match('//u', $string) === 1;
        }

        if (function_exists('iconv')) {
            $converted = @iconv($encoding, $encoding . '//IGNORE', $string);
            return $converted !== false && $converted === $string;
        }

        return true;
    }
}

if (!function_exists('mb_substitute_character')) {
    function mb_substitute_character($substitute_character = null)
    {
        static $current = 63;

        if ($substitute_character === null) {
            return $current;
        }

        $current = $substitute_character;
        return true;
    }
}

if (!function_exists('mb_ord')) {
    function mb_ord($string, $encoding = null)
    {
        $encoding = isop_exporter_mb_encoding($encoding);
        $converted = iconv($encoding, 'UCS-4BE', $string);
        $unpacked = unpack('N', $converted);
        return $unpacked ? $unpacked[1] : false;
    }
}

if (!function_exists('mb_chr')) {
    function mb_chr($code, $encoding = null)
    {
        $encoding = isop_exporter_mb_encoding($encoding);
        return iconv('UCS-4BE', $encoding, pack('N', $code));
    }
}
