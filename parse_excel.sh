#!/bin/bash

#Скрипт отбирает из нескольких файлов .xlsx заданные столбцы и пишет их в новый файл в заданном порядке

#Указываем каталог с файлами xicx
dir="/home/grmax/Загрузки/csv/"

#Конвертируем .xlsx в .txt в формате csv
for file in $(find "$dir" -type f -name "*.xlsx")
do

	fileout=$(echo $file | cut -d'.' -f1)

ssconvert -O 'separator=; format=raw' $file $fileout.txt  >/dev/null 2>&1

done

#Удаляем, на всякий случай файлик в который будем писать итог
rm out.csv >/dev/null 2>&1

#Читаем каждый файл .txt
for file in $(find "$dir" -type f -name "*.txt")
do

#переменная - одно из значений будущей таблицы, берем из названиия файла, можно опустить если не надо
filename=""

#Ищем интересующие нас столбцы, если столбец пуст добавляем сивол "0" в начало строки, преобразуем все символы в строчные, 
#для корректного поиска, удаляем символ переноса каретки, для точного определения столбца (если наименования столбцов схожи) используем поиск как в column7
column1=$(cat $file | head -1 | sed '/^;/ s/./0;/' | tr [:upper:] [:lower:] | tr ';' '\n' | nl |grep -w "column1" | tr -d " " | awk -F " " '{print $1}')
column2=$(cat $file | head -1 | sed '/^;/ s/./0;/' | tr [:upper:] [:lower:] | tr ';' '\n' | nl |grep -w "column2" | tr -d " " | awk -F " " '{print $1}')
column3=$(cat $file | head -1 | sed '/^;/ s/./0;/' | tr [:upper:] [:lower:] | tr ';' '\n' | nl |grep -w "column3" | tr -d " " | awk -F " " '{print $1}')
column4=$(cat $file | head -1 | sed '/^;/ s/./0;/' | tr [:upper:] [:lower:] | tr ';' '\n' | nl |grep -w "column4" | tr -d " " | awk -F " " '{print $1}')
column5=$(cat $file | head -1 | sed '/^;/ s/./0;/' | tr [:upper:] [:lower:] | tr ';' '\n' | nl |grep -w "column6" | tr -d " " | awk -F " " '{print $1}')
column6=$(cat $file | head -1 | sed '/^;/ s/./0;/' | tr [:upper:] [:lower:] | tr ';' '\n' | nl |grep -w "column6" | tr -d " " | awk -F " " '{print $1}')
column7=$(cat $file | head -1 | sed '/^;/ s/./0;/' | tr [:upper:] [:lower:] | tr ';' '\n' | nl |grep -w "\<column7\>" | tr -d " " | awk -F " " '{print $1}')


#Проверяем наличие всех столбцов в обрабатываемом файле если столбец отсутствует выдаем сообщение и выходим
if [ -z "${column1}" ];
then 
	echo "операция прервана в $file нет данных column1"
	exit
elif [ -z "${column2}" ];
then
	echo "операция прервана в $file нет данных column2"
	exit
elif [ -z "${column3}" ];
then
	echo "операция прервана в $file нет данных column3"
	exit
elif [ -z "${column4}" ];
then
	echo "операция прервана в $file нет данных column4"
	exit
elif [ -z "${column5}" ];
then
	echo "операция прервана в $file нет данных column5"
	exit
elif [ -z "${column6}" ];
then
	echo "операция прервана в $file нет данных column6"
	exit
elif [ -z "${column7}" ];
then
	echo "операция прервана в $file нет данных column7"
	exit
fi

#Формируем наименование файла (будет отдельным столбцом в таблице
if echo "$file" | tr [:upper:] [:lower:]   | grep "file_name1" > /dev/null
then
		filename="Файл 1"
	elif echo "$file" | tr [:upper:] [:lower:]   | grep "file_name2" > /dev/null
	then
		filename="Файл 2"
        elif echo "$file" | tr [:upper:] [:lower:]   | grep "file_name3" > /dev/null
        then
                filename="Файл 3"
#Если не получилось определить наименование, то пишем сообщение и выходим
	elif [ -z "${rayon}" ];
	then 
		echo "Не удалось определить наименование района по имени файла $file"
		exit
	fi

#Смотрим если файл не существует, то создаем его и вписываем первую строку с наименованиями требуемых столбцоа
if [ ! -f out.csv ]
then

cat $file | sed '/^;/ s/./0;/' | awk -v OFS=';' -F';' '{print \ 
        "filename", \
	$'$column1', \ 
       	$'$column2', \ 
       	$'$column3', \ 
       	$'$column4', \ 
       	$'$column5', \ 
       	$'$column6', \ 
       	$'$column7'}' | head -1 >out.csv

fi

#Построчно формируем таблицу со значениями из нужных столбцов
while IFS=";" read -r col1 col2  col3  col4  col5  col6  col7  
do

	echo "$filename;$col1;$col2;$col3;$col4;$col5;$col6;$col7"  >> out.csv

done < <(awk -v OFS=';' -F';' '{print \
        $'$column1', \
        $'$column2', \
        $'$column3', \
        $'$column4', \
        $'$column5', \
        $'$column6', \
        $'$column7'}'  $file | tail -n +2)
done

#Удаляем ненужные .txt
rm -f $dir.txt

#итоговый файл конвертируем в excel
	ssconvert  out.csv out.xlsx
#Удаляем временный итоговый файл
rm -f out.csv
