# <a name="requirements-element"></a>Элемент Requirements

Указывает минимальный набор элементов API JavaScript для Office ([набор требований](https://docs.microsoft.com/office/dev/add-ins/develop/office-versions-and-requirement-sets#specify-office-hosts-and-requirement-sets) и/или методов), необходимых для активации надстройки Office.

**Тип надстройки:** контентные и почтовые надстройки, надстройки области задач.

## <a name="syntax"></a>Синтаксис

```XML
<Requirements>
   ...
</Requirements>
```

## <a name="contained-in"></a>Содержащиеся в

[OfficeApp](officeapp.md)

## <a name="can-contain"></a>Может содержать

|**Элемент**|**Контентная надстройка**|**Почта**|**Область задач**|
|:-----|:-----|:-----|:-----|
|[Sets](sets.md)|x|x|x|
|[Методы](methods.md)|x||x|

## <a name="remarks"></a>Замечания

Дополнительные сведения о наборах требований [версии Office и требования наборов](https://docs.microsoft.com/office/dev/add-ins/develop/office-versions-and-requirement-sets)см.

