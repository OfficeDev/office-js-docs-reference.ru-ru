# <a name="allformfactors-element"></a><span data-ttu-id="f049f-101">Элемент AllFormFactors</span><span class="sxs-lookup"><span data-stu-id="f049f-101">AllFormFactors element</span></span>

<span data-ttu-id="f049f-102">Указывает параметры всех форм-факторов для надстройки.</span><span class="sxs-lookup"><span data-stu-id="f049f-102">Specifies the settings for an add-in for all form factors.</span></span> <span data-ttu-id="f049f-103">На данный момент только компонент с помощью **AllFormFactors** — пользовательских функций.</span><span class="sxs-lookup"><span data-stu-id="f049f-103">Currently, the only feature using **AllFormFactors** is custom functions.</span></span> <span data-ttu-id="f049f-104">**AllFormFactors** является обязательным элементом при использовании пользовательских функций.</span><span class="sxs-lookup"><span data-stu-id="f049f-104">**AllFormFactors** is a required element when using custom functions.</span></span>

## <a name="child-elements"></a><span data-ttu-id="f049f-105">Дочерние элементы</span><span class="sxs-lookup"><span data-stu-id="f049f-105">Child elements</span></span>

|  <span data-ttu-id="f049f-106">Элемент</span><span class="sxs-lookup"><span data-stu-id="f049f-106">Element</span></span> |  <span data-ttu-id="f049f-107">Обязательный</span><span class="sxs-lookup"><span data-stu-id="f049f-107">Required</span></span>  |  <span data-ttu-id="f049f-108">Описание</span><span class="sxs-lookup"><span data-stu-id="f049f-108">Description</span></span>  |
|:-----|:-----|:-----|
|  [<span data-ttu-id="f049f-109">ExtensionPoint</span><span class="sxs-lookup"><span data-stu-id="f049f-109">ExtensionPoint</span></span>](extensionpoint.md) |  <span data-ttu-id="f049f-110">Да</span><span class="sxs-lookup"><span data-stu-id="f049f-110">Yes</span></span> |  <span data-ttu-id="f049f-111">Определяет, где предоставляются функции надстройки.</span><span class="sxs-lookup"><span data-stu-id="f049f-111">Defines where an add-in exposes functionality.</span></span> |

## <a name="allformfactors-example"></a><span data-ttu-id="f049f-112">Пример использования AllFormFactors</span><span class="sxs-lookup"><span data-stu-id="f049f-112">AllFormFactors example</span></span>

```xml
<Hosts>
    <Host xsi:type="Workbook">
        <AllFormFactors>
            <ExtensionPoint xsi:type="CustomFunctions">
                    <!-- Information on this extension point -->
            </ExtensionPoint>
        </AllFormFactors>
    </Host>
</Hosts>
```
