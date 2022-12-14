# javaDocxReplace

## Основные библиотеки:

1. [Docx4](https://github.com/phip1611/docx4j-search-and-replace-util)
2. com.spire.doc

## Как работает код

Если хотите, что бы в вашем шаблоне менялись определенные переменные на ваши текстовые значения, то используем метод:
```DocumentReplacment.documentReplace(String nameTemplate, String[][] stringToReplaceAndToReplacment, String outputFileName)```.

Где:

`nameTemplate` - относительный путь до шаблона файла

`stringToReplaceAndToReplacment` - массив для с заменяемыми переменными на их значения

`outputFileName` - выходной файл, будет появляется в output
Этот метод в конце выдает путь output файла. 

Пример кода:  
~~~
String [] [] stringToReplaceAndToReplacment = new String[][]{
    {"fio", "Матвеев Дмитрий Владимирович"},
    {"special", "Информационные системы"},
    {"fioTo", "Олег Иванович"},
};

String template = "Заявление.docx";

template = DocumentReplacment.documentReplace("template.docx", stringToReplaceAndToReplacment, template);
~~~
Если хотите поменять переменную на изображение, то используйте метод:
```DocumentReplacment.replaceTextWithImage(String inputPath, String stringToReplace, String imagePath)```, где: 

`inputPath` - путь до файла; 

`stringToReplace` - название для переменной внутри шаблона;

`imagePath` - путь до картинки

Если в первом методе создается собственный файл из шаблона, то во втором случае метод заменяет выбранный файл.