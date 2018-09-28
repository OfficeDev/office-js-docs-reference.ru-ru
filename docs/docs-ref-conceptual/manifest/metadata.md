# <a name="metadata-element"></a><span data-ttu-id="aa922-101">Элемент Metadata</span><span class="sxs-lookup"><span data-stu-id="aa922-101">Metadata element</span></span>

<span data-ttu-id="aa922-102">Задает параметры метаданных, используемых пользовательских функций в Excel.</span><span class="sxs-lookup"><span data-stu-id="aa922-102">Defines the metadata settings used by a custom function in Excel.</span></span>

## <a name="attributes"></a><span data-ttu-id="aa922-103">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="aa922-103">Attributes</span></span>

<span data-ttu-id="aa922-104">Нет</span><span class="sxs-lookup"><span data-stu-id="aa922-104">None</span></span>

## <a name="child-elements"></a><span data-ttu-id="aa922-105">Дочерние элементы</span><span class="sxs-lookup"><span data-stu-id="aa922-105">Child elements</span></span>

|  <span data-ttu-id="aa922-106">Элемент</span><span class="sxs-lookup"><span data-stu-id="aa922-106">Element</span></span>  |  <span data-ttu-id="aa922-107">Обязательный</span><span class="sxs-lookup"><span data-stu-id="aa922-107">Required</span></span>  |  <span data-ttu-id="aa922-108">Описание</span><span class="sxs-lookup"><span data-stu-id="aa922-108">Description</span></span>  |
|:-----|:-----|:-----|
|  [<span data-ttu-id="aa922-109">SourceLocation</span><span class="sxs-lookup"><span data-stu-id="aa922-109">SourceLocation</span></span>](customfunctionssourcelocation.md)  |  <span data-ttu-id="aa922-110">Да</span><span class="sxs-lookup"><span data-stu-id="aa922-110">Yes</span></span>  | <span data-ttu-id="aa922-111">Строка с идентификатором ресурсов из файла JSON, используемого пользовательских функций.</span><span class="sxs-lookup"><span data-stu-id="aa922-111">String with the resource id of the JSON file used by custom functions.</span></span> |

## <a name="example"></a><span data-ttu-id="aa922-112">Пример</span><span class="sxs-lookup"><span data-stu-id="aa922-112">Example</span></span>

```xml
<Metadata>
    <SourceLocation resid="JSON-URL" />
</Metadata>
```
