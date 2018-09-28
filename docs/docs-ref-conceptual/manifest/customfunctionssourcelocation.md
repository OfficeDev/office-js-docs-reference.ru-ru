# <a name="sourcelocation-element"></a><span data-ttu-id="5fa23-101">Элемент SourceLocation</span><span class="sxs-lookup"><span data-stu-id="5fa23-101">SourceLocation element</span></span>

<span data-ttu-id="5fa23-102">Определяет расположение ресурса, который необходим для элементов Script или Page, используемых пользовательскими функциями в Excel.</span><span class="sxs-lookup"><span data-stu-id="5fa23-102">Defines the location of a resource needed by the Script or Page elements used by custom functions in Excel.</span></span>

## <a name="attributes"></a><span data-ttu-id="5fa23-103">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="5fa23-103">Attributes</span></span>

| <span data-ttu-id="5fa23-104">**Атрибут**</span><span class="sxs-lookup"><span data-stu-id="5fa23-104">**Attribute**</span></span> | <span data-ttu-id="5fa23-105">**Обязательное**</span><span class="sxs-lookup"><span data-stu-id="5fa23-105">**Required**</span></span> | <span data-ttu-id="5fa23-106">**Описание**</span><span class="sxs-lookup"><span data-stu-id="5fa23-106">**Description**</span></span>                                                                      |
|---------------|--------------|--------------------------------------------------------------------------------------|
| <span data-ttu-id="5fa23-107">resid</span><span class="sxs-lookup"><span data-stu-id="5fa23-107">resid</span></span>         | <span data-ttu-id="5fa23-108">Да</span><span class="sxs-lookup"><span data-stu-id="5fa23-108">Yes</span></span>          | <span data-ttu-id="5fa23-109">Имя ресурса URL-адреса, определенного в разделе &lt;Ресурсы&gt; в манифесте.</span><span class="sxs-lookup"><span data-stu-id="5fa23-109">The name of a URL resource defined in the &lt;Resources&gt; section of the manifest.</span></span> |

## <a name="child-elements"></a><span data-ttu-id="5fa23-110">Дочерние элементы</span><span class="sxs-lookup"><span data-stu-id="5fa23-110">Child elements</span></span>

<span data-ttu-id="5fa23-111">Нет</span><span class="sxs-lookup"><span data-stu-id="5fa23-111">None</span></span>

## <a name="example"></a><span data-ttu-id="5fa23-112">Пример</span><span class="sxs-lookup"><span data-stu-id="5fa23-112">Example</span></span>

```xml
<SourceLocation resid="pageURL"/>
```