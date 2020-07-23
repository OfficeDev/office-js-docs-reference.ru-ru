# <a name="office-javascript-api-reference"></a><span data-ttu-id="0a280-101">Справочные материалы по API JavaScript для Office</span><span class="sxs-lookup"><span data-stu-id="0a280-101">Office JavaScript API Reference</span></span>

<span data-ttu-id="0a280-102">Добро пожаловать в репозиторий справочной документации по API JavaScript для Office.</span><span class="sxs-lookup"><span data-stu-id="0a280-102">Welcome to the Office JavaScript API Reference documentation repository.</span></span> <span data-ttu-id="0a280-103">Рекомендуем просматривать эти материалы на сайте [docs.microsoft.com](https://docs.microsoft.com/javascript/api/overview/office).</span><span class="sxs-lookup"><span data-stu-id="0a280-103">For the best experience, we recommend you view this content on [docs.microsoft.com](https://docs.microsoft.com/javascript/api/overview/office).</span></span>

> <span data-ttu-id="0a280-104">**Примечание**: исходные файлы документации можно найти в статье Общие сведения об API JavaScript для Office, краткие руководства, учебные пособия и пошаговые руководства в репозитории [OfficeDev/Office-JS-Documentation-PR](https://github.com/OfficeDev/office-js-docs-pr) .</span><span class="sxs-lookup"><span data-stu-id="0a280-104">**Note**: You can find the documentation source files for Office JavaScript API concepts, quick starts, tutorials, and how-to guides in the [OfficeDev/office-js-docs-pr](https://github.com/OfficeDev/office-js-docs-pr) repository.</span></span>

## <a name="give-us-your-feedback"></a><span data-ttu-id="0a280-105">Оставьте свой отзыв</span><span class="sxs-lookup"><span data-stu-id="0a280-105">Give us your feedback</span></span>

<span data-ttu-id="0a280-106">Ваше мнение важно для нас.</span><span class="sxs-lookup"><span data-stu-id="0a280-106">Your feedback is important to us.</span></span>

* <span data-ttu-id="0a280-107">Чтобы задать нам вопрос или сообщить о проблемах с документацией, [оставьте сообщение](https://github.com/OfficeDev/office-js-docs-reference/issues) на вкладке этого репозитория.</span><span class="sxs-lookup"><span data-stu-id="0a280-107">To let us know about any questions or issues you find in the docs, [submit an issue](https://github.com/OfficeDev/office-js-docs-reference/issues) in this repository.</span></span> <span data-ttu-id="0a280-108">Убедитесь, что вы задаете номер версии + номер сборки клиента, который вы используете, и при необходимости предоставляет процедуры воспроизведения, выходные данные консоли и сообщения об ошибках.</span><span class="sxs-lookup"><span data-stu-id="0a280-108">Make sure you state the version + build number of the client you are using, and provide repro steps, console output, and error messages, as appropriate.</span></span>

* <span data-ttu-id="0a280-109">Мы также будем рады Вашим вкладам в эту документацию.</span><span class="sxs-lookup"><span data-stu-id="0a280-109">We also welcome your contributions to this documentation.</span></span> <span data-ttu-id="0a280-110">Чтобы внести изменения, разработайте этот репозиторий, обновите файлы по мере необходимости и отправьте запрос на включение внесенных изменений.</span><span class="sxs-lookup"><span data-stu-id="0a280-110">To contribute, fork this repository, update the files as you deem necessary, and submit a pull request with your proposed changes.</span></span> <span data-ttu-id="0a280-111">Дополнительные сведения см [в статье участие в этой документации](Contributing.md).</span><span class="sxs-lookup"><span data-stu-id="0a280-111">For details, see [Contribute to this documentation](Contributing.md).</span></span>

    > <span data-ttu-id="0a280-112">**Важно!** не изменяйте файлы в папке [/Докс/Докс-реф-аутожен](https://github.com/OfficeDev/office-js-docs-reference/tree/master/docs/docs-ref-autogen) этого репозитория.</span><span class="sxs-lookup"><span data-stu-id="0a280-112">**IMPORTANT**: Do not modify files within the [/docs/docs-ref-autogen](https://github.com/OfficeDev/office-js-docs-reference/tree/master/docs/docs-ref-autogen) folder of this repository.</span></span> <span data-ttu-id="0a280-113">Все файлы в этой папке создаются автоматически, поэтому их невозможно обновлять с помощью запроса на включение внесенных изменений.</span><span class="sxs-lookup"><span data-stu-id="0a280-113">All of the files in that folder are autogenerated, so it is not possible to update them via pull request.</span></span> <span data-ttu-id="0a280-114">Чтобы запросить изменение файлов в папке [/Докс/Докс-реф-аутожен](https://github.com/OfficeDev/office-js-docs-reference/tree/master/docs/docs-ref-autogen) , [отправьте сообщение об ошибке](https://github.com/OfficeDev/office-js-docs-reference/issues) в этом репозитории.</span><span class="sxs-lookup"><span data-stu-id="0a280-114">To request a change to any of the files in the [/docs/docs-ref-autogen](https://github.com/OfficeDev/office-js-docs-reference/tree/master/docs/docs-ref-autogen) folder, please [submit an issue](https://github.com/OfficeDev/office-js-docs-reference/issues) in this repository.</span></span> <span data-ttu-id="0a280-115">Подробнее о том, как инструментарий в этом репозитории можно прочитать [здесь](https://github.com/OfficeDev/office-js-docs-reference/blob/master/DocumentationToolingNotes.md).</span><span class="sxs-lookup"><span data-stu-id="0a280-115">You can read more about how the tooling in this repository [here](https://github.com/OfficeDev/office-js-docs-reference/blob/master/DocumentationToolingNotes.md).</span></span>

* <span data-ttu-id="0a280-116">Чтобы рассказать о своих впечатлениях от использования файлов и пожеланиях насчет будущих версий, примеров кода и т. д., поделитесь своими мыслями на [сайте UserVoice платформы разработки для Office](https://officespdev.uservoice.com/).</span><span class="sxs-lookup"><span data-stu-id="0a280-116">To let us know about your programming experience, what you would like to see in future versions, code samples, and so on, enter your suggestions and ideas at [Office Developer Platform UserVoice](https://officespdev.uservoice.com/).</span></span>

## <a name="microsoft-open-source-code-of-conduct"></a><span data-ttu-id="0a280-117">Правила поведения Майкрософт, касающиеся обращения с открытым кодом</span><span class="sxs-lookup"><span data-stu-id="0a280-117">Microsoft Open Source Code of Conduct</span></span>

<span data-ttu-id="0a280-118">Этот проект соответствует [правилам поведения Майкрософт, касающимся обращения с открытым кодом](https://opensource.microsoft.com/codeofconduct/).</span><span class="sxs-lookup"><span data-stu-id="0a280-118">This project has adopted the [Microsoft Open Source Code of Conduct](https://opensource.microsoft.com/codeofconduct/).</span></span>
<span data-ttu-id="0a280-119">Для получения дополнительных сведений обратитесь к разделу " [проведение вопросов](https://opensource.microsoft.com/codeofconduct/faq/)" или свяжитесь с [opencode@microsoft.com](mailto:opencode@microsoft.com) с дополнительными вопросами или комментариями.</span><span class="sxs-lookup"><span data-stu-id="0a280-119">For more information, see the [Code of Conduct FAQ](https://opensource.microsoft.com/codeofconduct/faq/), or contact [opencode@microsoft.com](mailto:opencode@microsoft.com) with any additional questions or comments.</span></span>
