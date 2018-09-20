# <a name="webapplicationinfo-element"></a><span data-ttu-id="992d3-101">Элемент WebApplicationInfo</span><span class="sxs-lookup"><span data-stu-id="992d3-101">WebApplicationInfo element</span></span>

<span data-ttu-id="992d3-102">Поддерживает единый вход в надстройках Office. Этот элемент содержит сведения для надстройки в качестве следующего:</span><span class="sxs-lookup"><span data-stu-id="992d3-102">Supports single sign-on (SSO) in Office Add-ins. This element contains information for the add-in as both:</span></span>

- <span data-ttu-id="992d3-103">*Ресурс* OAuth 2.0, для которого могут потребоваться разрешения ведущему приложению Office.</span><span class="sxs-lookup"><span data-stu-id="992d3-103">An OAuth 2.0 *resource* to which the Office host application might need permissions.</span></span>
- <span data-ttu-id="992d3-104">*Клиент* OAuth 2.0, которому могут потребоваться разрешения для Microsoft Graph.</span><span class="sxs-lookup"><span data-stu-id="992d3-104">An OAuth 2.0 *client* that might need permissions to Microsoft Graph.</span></span>

<span data-ttu-id="992d3-105">**WebApplicationInfo** — дочерний элемент элемента [VersionOverrides](versionoverrides.md) в манифесте.</span><span class="sxs-lookup"><span data-stu-id="992d3-105">**WebApplicationInfo** is a child element of the [VersionOverrides](versionoverrides.md) element in the manifest.</span></span>  

## <a name="child-elements"></a><span data-ttu-id="992d3-106">Дочерние элементы</span><span class="sxs-lookup"><span data-stu-id="992d3-106">Child elements</span></span>

|  <span data-ttu-id="992d3-107">Элемент</span><span class="sxs-lookup"><span data-stu-id="992d3-107">Element</span></span> |  <span data-ttu-id="992d3-108">Обязательный</span><span class="sxs-lookup"><span data-stu-id="992d3-108">Required</span></span>  |  <span data-ttu-id="992d3-109">Описание</span><span class="sxs-lookup"><span data-stu-id="992d3-109">Description</span></span>  |
|:-----|:-----|:-----|
|  <span data-ttu-id="992d3-110">**Id**</span><span class="sxs-lookup"><span data-stu-id="992d3-110">**Id**</span></span>    |  <span data-ttu-id="992d3-111">Да</span><span class="sxs-lookup"><span data-stu-id="992d3-111">Yes</span></span>   |  <span data-ttu-id="992d3-112">**Идентификатор** связанной с надстройкой службы, зарегистрированный в конечной точке Azure Active Directory 2.0.</span><span class="sxs-lookup"><span data-stu-id="992d3-112">The **Application Id** of the add-in's associated service as registered in the Azure Active Directory v 2.0 endpoint.</span></span>|
|  <span data-ttu-id="992d3-113">**Resource**</span><span class="sxs-lookup"><span data-stu-id="992d3-113">**Resource**</span></span>  |  <span data-ttu-id="992d3-114">Да</span><span class="sxs-lookup"><span data-stu-id="992d3-114">Yes</span></span>   |  <span data-ttu-id="992d3-115">Указывает **URI идентификатора** надстройки, зарегистрированный в конечной точке Azure Active Directory 2.0.</span><span class="sxs-lookup"><span data-stu-id="992d3-115">Specifies the **Application ID URI** of the add-in as registered in the Azure Active Directory v 2.0 endpoint.</span></span>|
|  [<span data-ttu-id="992d3-116">Scopes</span><span class="sxs-lookup"><span data-stu-id="992d3-116">Scopes</span></span>](scopes.md)                |  <span data-ttu-id="992d3-117">Нет</span><span class="sxs-lookup"><span data-stu-id="992d3-117">No</span></span>  |  <span data-ttu-id="992d3-118">Указывает разрешения, необходимые надстройке для работы с Microsoft Graph.</span><span class="sxs-lookup"><span data-stu-id="992d3-118">Specifies the permissions that the add-in needs to Microsoft Graph.</span></span>  |

> [!NOTE] 
> <span data-ttu-id="992d3-119">В настоящее время это необходимо, что ресурс надстройки соответствует его узла.</span><span class="sxs-lookup"><span data-stu-id="992d3-119">Currently, it's necessary that your add-in's Resource matches its Host.</span></span> <span data-ttu-id="992d3-120">Office запрашивает маркер для надстройки, только если может подтвердить право собственности. В настоящее время для этого необходимо, чтобы надстройка размещалась под полным доменным именем ресурса.</span><span class="sxs-lookup"><span data-stu-id="992d3-120">Office will not request a Token for an add-in unless it can prove ownership, and today this is done by hosting the add-in under the Resource's fully-qualified domain name.</span></span>

## <a name="webapplicationinfo-example"></a><span data-ttu-id="992d3-121">Пример WebApplicationInfo</span><span class="sxs-lookup"><span data-stu-id="992d3-121">WebApplicationInfo example</span></span>

```xml
<OfficeApp>
...
  <VersionOverrides xmlns="http://schemas.microsoft.com/office/mailappversionoverrides" xsi:type="VersionOverridesV1_0">
    ...
    <WebApplicationInfo>
      <Id>12345678-abcd-1234-efab-123456789abc</Id>
      <Resource>api://myDomain.com/12345678-abcd-1234-efab-123456789abc<Resource>
      <Scopes>
        <Scope>Files.Read.All</Scope>
        <Scope>offline_access</Scope>
        <Scope>openid</Scope>
        <Scope>profile</Scope>        
      </Scopes>
    </WebApplicationInfo>
  </VersionOverrides>
...
</OfficeApp>
```
