# <a name="script-element"></a><span data-ttu-id="fc708-101">Элемент Script</span><span class="sxs-lookup"><span data-stu-id="fc708-101">Script element</span></span>

<span data-ttu-id="fc708-102">Определяет параметры сценариев, используемых пользовательской функцией в Excel.</span><span class="sxs-lookup"><span data-stu-id="fc708-102">Defines script settings used by a custom function in Excel.</span></span>

## <a name="attributes"></a><span data-ttu-id="fc708-103">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="fc708-103">Attributes</span></span>

<span data-ttu-id="fc708-104">Нет</span><span class="sxs-lookup"><span data-stu-id="fc708-104">None</span></span>

## <a name="child-elements"></a><span data-ttu-id="fc708-105">Дочерние элементы</span><span class="sxs-lookup"><span data-stu-id="fc708-105">Child elements</span></span>

|<span data-ttu-id="fc708-106">Элементы</span><span class="sxs-lookup"><span data-stu-id="fc708-106">Elements</span></span>  |  <span data-ttu-id="fc708-107">Обязательный</span><span class="sxs-lookup"><span data-stu-id="fc708-107">Required</span></span>  |  <span data-ttu-id="fc708-108">Описание</span><span class="sxs-lookup"><span data-stu-id="fc708-108">Description</span></span>  |
|:-----|:-----|:-----|
|  [<span data-ttu-id="fc708-109">SourceLocation</span><span class="sxs-lookup"><span data-stu-id="fc708-109">SourceLocation</span></span>](customfunctionssourcelocation.md)  |  <span data-ttu-id="fc708-110">Да</span><span class="sxs-lookup"><span data-stu-id="fc708-110">Yes</span></span>  | <span data-ttu-id="fc708-111">Строка с идентификатором ресурса файла JavaScript, используемого пользовательскими функциями.</span><span class="sxs-lookup"><span data-stu-id="fc708-111">String with resource id of the JavaScript file used by custom functions.</span></span>|

## <a name="example"></a><span data-ttu-id="fc708-112">Пример</span><span class="sxs-lookup"><span data-stu-id="fc708-112">Example</span></span>

```xml
<Script>
    <SourceLocation resid="scriptURL" />
    <!-- The Script element is not used in the Developer Preview. -->
</Script>
```
