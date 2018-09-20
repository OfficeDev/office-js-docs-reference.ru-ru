# <a name="desktopformfactor-element"></a><span data-ttu-id="22cfb-101">Элемент DesktopFormFactor</span><span class="sxs-lookup"><span data-stu-id="22cfb-101">DesktopFormFactor element</span></span>

<span data-ttu-id="22cfb-p101">Указывает параметры для надстройки классического форм-фактора. Классический форм-фактор включает Office для Windows, Office для Mac и Office Online. Он содержит все сведения о надстройке для классического форм-фактора, кроме узла **Resources**.</span><span class="sxs-lookup"><span data-stu-id="22cfb-p101">Specifies the settings for an add-in for the desktop form factor. The desktop form factor includes Office for Windows, Office for Mac, and Office Online. It contains all the add-in information for the desktop form factor except for the  **Resources** node.</span></span>

<span data-ttu-id="22cfb-p102">В каждом определении DesktopFormFactor есть элемент **FunctionFile**, а также один или несколько элементов **ExtensionPoint**. Дополнительные сведения см. в статьях [Элемент FunctionFile](functionfile.md) и [Элемент ExtensionPoint](extensionpoint.md).</span><span class="sxs-lookup"><span data-stu-id="22cfb-p102">Each DesktopFormFactor definition contains the  **FunctionFile** element and one or more **ExtensionPoint** elements. For more information, see [FunctionFile element](functionfile.md) and [ExtensionPoint element](extensionpoint.md).</span></span> 

## <a name="child-elements"></a><span data-ttu-id="22cfb-107">Дочерние элементы</span><span class="sxs-lookup"><span data-stu-id="22cfb-107">Child elements</span></span>

| <span data-ttu-id="22cfb-108">Элемент</span><span class="sxs-lookup"><span data-stu-id="22cfb-108">Element</span></span>                               | <span data-ttu-id="22cfb-109">Обязательный</span><span class="sxs-lookup"><span data-stu-id="22cfb-109">Required</span></span> | <span data-ttu-id="22cfb-110">Описание</span><span class="sxs-lookup"><span data-stu-id="22cfb-110">Description</span></span>  |
|:--------------------------------------|:--------:|:-------------|
| [<span data-ttu-id="22cfb-111">ExtensionPoint</span><span class="sxs-lookup"><span data-stu-id="22cfb-111">ExtensionPoint</span></span>](extensionpoint.md) | <span data-ttu-id="22cfb-112">Да</span><span class="sxs-lookup"><span data-stu-id="22cfb-112">Yes</span></span>      | <span data-ttu-id="22cfb-113">Определяет, где предоставляются функции надстройки.</span><span class="sxs-lookup"><span data-stu-id="22cfb-113">Defines where an add-in exposes functionality.</span></span> |
| [<span data-ttu-id="22cfb-114">FunctionFile</span><span class="sxs-lookup"><span data-stu-id="22cfb-114">FunctionFile</span></span>](functionfile.md)     | <span data-ttu-id="22cfb-115">Да</span><span class="sxs-lookup"><span data-stu-id="22cfb-115">Yes</span></span>      | <span data-ttu-id="22cfb-116">URL-адрес файла, который содержит функции JavaScript.</span><span class="sxs-lookup"><span data-stu-id="22cfb-116">A URL to a file that contains JavaScript functions.</span></span>|
| [<span data-ttu-id="22cfb-117">GetStarted</span><span class="sxs-lookup"><span data-stu-id="22cfb-117">GetStarted</span></span>](getstarted.md)         | <span data-ttu-id="22cfb-118">Нет</span><span class="sxs-lookup"><span data-stu-id="22cfb-118">No</span></span>       | <span data-ttu-id="22cfb-119">Определяет выноску, которая отображается при установке надстройки в ведущих приложениях Word, Excel и PowerPoint.</span><span class="sxs-lookup"><span data-stu-id="22cfb-119">Defines the callout that appears when installing the add-in in Word, Excel, or PowerPoint hosts.</span></span> |

## <a name="desktopformfactor-example"></a><span data-ttu-id="22cfb-120">Пример DesktopFormFactor</span><span class="sxs-lookup"><span data-stu-id="22cfb-120">DesktopFormFactor example</span></span>

```xml
...
<Hosts>
  <Host xsi:type="Presentation">
    <DesktopFormFactor>
      <FunctionFile resid="residDesktopFuncUrl" />
      <GetStarted>
        <!-- GetStarted callout -->
      </GetStarted>
      <ExtensionPoint xsi:type="PrimaryCommandSurface">
        <!-- information on this extension point -->
      </ExtensionPoint> 
      <!-- possibly more ExtensionPoint elements -->
    </DesktopFormFactor>
  </Host>
</Hosts>
...
```
