# <a name="set-element"></a>Элемент Set

Указывает набор требований из API JavaScript для Office, необходимый для активации надстройки Office.

**Тип надстройки:** контентные и почтовые надстройки, надстройки области задач.

## <a name="syntax"></a>Синтаксис

```XML
<Set Name="string" MinVersion="n .n">
```

## <a name="contained-in"></a>Содержащиеся в

[Sets](sets.md)

## <a name="attributes"></a>Атрибуты

|**Атрибут**|**Тип**|**Обязательный**|**Описание**|
|:-----|:-----|:-----|:-----|
|Имя|string|Обязательный|Имя [набора требований](https://docs.microsoft.com/office/dev/add-ins/develop/office-versions-and-requirement-sets).|
|MinVersion|string|необязательный|Указывает минимальную версию набора API, необходимую надстройке. Переопределяет значение **DefaultMinVersion**, если оно указано в родительском элементе [Sets](sets.md).|

## <a name="remarks"></a>Замечания

Дополнительные сведения о наборах требований [версии Office и требования наборов](https://docs.microsoft.com/office/dev/add-ins/develop/office-versions-and-requirement-sets)см.

Дополнительные сведения об атрибуте **MinVersion** элемента **Set** и атрибуте **DefaultMinVersion** элемента **Sets** см. в статье [Указание элемента Requirements в манифесте](https://docs.microsoft.com/office/dev/add-ins/develop/specify-office-hosts-and-api-requirements#set-the-requirements-element-in-the-manifest).

> [!IMPORTANT] 
> Для надстроек почты существует только один `"Mailbox"` набору требований. Этот набор требований содержит всей подмножество API, поддерживаемые в надстройках почты для Outlook, и необходимо указать `"Mailbox"` требований в почты надстроек в его манифесте (это не необязательно как в случае содержимого и задач надстроек области). Кроме того невозможно объявить поддержку для отдельных методов в надстройках почты.
