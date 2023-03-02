<?php

namespace Ayflying\Excel2php;

use PhpOffice\PhpSpreadsheet\IOFactory;

class Load
{

    /**
     * 文件转数据表
     * @param string $file
     * @param int $fidelKey 表中key的字段
     * @param int $startRow 表内容从第几行开始
     * @return array
     */
    public static function getFile(string $file, int $fidelKey = 2, int $startRow = 3): array
    {
        $arr = self::excel2array($file);
        $list = [];
        //循环所有页签
        foreach ($arr as $title => $sheel) {
            $keys = $sheel[$fidelKey - 1] ?? null;
            //不存在页头的数据跳过
            if (empty($keys)) {
                continue;
            }
            $data = [];
            //循环所有行
            foreach ($sheel as $k => $val) {
                //当前行未开始
                if ($k < $startRow) {
                    continue;
                }
                $temp = [];
                foreach ($keys as $k2 => $val2) {
                    //跳过空key
                    if (empty($val2)) {
                        continue;
                    }
                    $temp[$val2] = $val[$k2];
                }
                $data[] = $temp;
            }
            $list[$title] = $data;
        }
        return $list;
    }

    /**
     * 整个目录表转数组
     * @param string $path 需要运行的文件夹
     * @param int $fidelKey 表中key的字段
     * @param int $startRow 表内容从第几行开始
     * @return array
     */
    public static function getPath(string $path, int $fidelKey = 2, int $startRow = 3): array
    {
        $arrFiles = glob($path . '/*.xlsx');
        $list = [];
        //循环总文件
        foreach ($arrFiles as $file) {
            $res = self::getFile($file, $fidelKey, $startRow);
            $list = array_merge($list, $res);
        }
        return $list;
    }

    /**
     * 表格转数组
     * @param string $file excel文件位置
     * @return array
     */
    public static function excel2array(string $file): array
    {
        if (!is_file($file)) {
            return [];
        }
        //获取所有标签
        $Sheets = IOFactory::load($file)->getAllSheets() ?? [];
        $list = [];
        foreach ($Sheets as $Sheet) {
            //获取工作表页签名称
            $title = $Sheet->getTitle();
            //所有数据转字段
            $res = $Sheet->getCellCollection()->getParent()->toArray();
            $list[$title] = $res;
        }
        return $list;
    }
}