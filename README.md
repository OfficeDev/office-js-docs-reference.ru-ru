# <a name="office-javascript-api-reference"></a><span data-ttu-id="80b00-101">Ссылка на API JavaScript Office</span><span class="sxs-lookup"><span data-stu-id="80b00-101">Office JavaScript API Reference</span></span>

<span data-ttu-id="80b00-102">Добро пожаловать в репозиторий справочной документации по API Office JavaScript.</span><span class="sxs-lookup"><span data-stu-id="80b00-102">Welcome to the Office JavaScript API Reference documentation repository.</span></span> <span data-ttu-id="80b00-103">Рекомендуем просматривать эти материалы на сайте [docs.microsoft.com](https://docs.microsoft.com/javascript/api/overview/office).</span><span class="sxs-lookup"><span data-stu-id="80b00-103">For the best experience, we recommend you view this content on [docs.microsoft.com](https://docs.microsoft.com/javascript/api/overview/office).</span></span>

> <span data-ttu-id="80b00-104">**Примечание.** В репозитории [OfficeDev/office-js-docs-pr](https://github.com/OfficeDev/office-js-docs-pr) можно найти исходные файлы документации для концепций API Office JavaScript, быстрые старты, учебники и инструкции по управлению.</span><span class="sxs-lookup"><span data-stu-id="80b00-104">**Note**: You can find the documentation source files for Office JavaScript API concepts, quick starts, tutorials, and how-to guides in the [OfficeDev/office-js-docs-pr](https://github.com/OfficeDev/office-js-docs-pr) repository.</span></span>

## <a name="give-us-your-feedback"></a><span data-ttu-id="80b00-105">Оставьте свой отзыв</span><span class="sxs-lookup"><span data-stu-id="80b00-105">Give us your feedback</span></span>

<span data-ttu-id="80b00-106">Ваше мнение важно для нас.</span><span class="sxs-lookup"><span data-stu-id="80b00-106">Your feedback is important to us.</span></span>

* <span data-ttu-id="80b00-107">Чтобы задать нам вопрос или сообщить о проблемах с документацией, [оставьте сообщение](https://github.com/OfficeDev/office-js-docs-reference/issues) на вкладке этого репозитория.</span><span class="sxs-lookup"><span data-stu-id="80b00-107">To let us know about any questions or issues you find in the docs, [submit an issue](https://github.com/OfficeDev/office-js-docs-reference/issues) in this repository.</span></span> <span data-ttu-id="80b00-108">Убедитесь, что вы задайте версию + число сборки клиента, используемого вами, и при необходимости предостерегаете этапы перепроцесирования, выход консоли и сообщения об ошибках.</span><span class="sxs-lookup"><span data-stu-id="80b00-108">Make sure you state the version + build number of the client you are using, and provide repro steps, console output, and error messages, as appropriate.</span></span>

* <span data-ttu-id="80b00-109">Мы также приветствуем ваши вклады в эту документацию.</span><span class="sxs-lookup"><span data-stu-id="80b00-109">We also welcome your contributions to this documentation.</span></span> <span data-ttu-id="80b00-110">Чтобы внести вклад в этот репозиторий, обновите файлы по мере необходимости и отправьте запрос на тягу с предложенными изменениями.</span><span class="sxs-lookup"><span data-stu-id="80b00-110">To contribute, fork this repository, update the files as you deem necessary, and submit a pull request with your proposed changes.</span></span> <span data-ttu-id="80b00-111">Подробные сведения см. [в материале Contribute to this documentation.](Contributing.md)</span><span class="sxs-lookup"><span data-stu-id="80b00-111">For details, see [Contribute to this documentation](Contributing.md).</span></span>

    > <span data-ttu-id="80b00-112">**ВАЖНО.** Не изменять файлы в [папке /docs/docs-ref-autogen](https://github.com/OfficeDev/office-js-docs-reference/tree/master/docs/docs-ref-autogen) этого репозитория.</span><span class="sxs-lookup"><span data-stu-id="80b00-112">**IMPORTANT**: Do not modify files within the [/docs/docs-ref-autogen](https://github.com/OfficeDev/office-js-docs-reference/tree/master/docs/docs-ref-autogen) folder of this repository.</span></span> <span data-ttu-id="80b00-113">Все файлы в этой папке автогенерированы, поэтому их невозможно обновить с помощью запроса на вытягивать.</span><span class="sxs-lookup"><span data-stu-id="80b00-113">All of the files in that folder are autogenerated, so it is not possible to update them via pull request.</span></span> <span data-ttu-id="80b00-114">Чтобы запросить изменение любого из файлов в [папке /docs/docs-ref-autogen,](https://github.com/OfficeDev/office-js-docs-reference/tree/master/docs/docs-ref-autogen) отправьте проблему [в](https://github.com/OfficeDev/office-js-docs-reference/issues) этом репозитории.</span><span class="sxs-lookup"><span data-stu-id="80b00-114">To request a change to any of the files in the [/docs/docs-ref-autogen](https://github.com/OfficeDev/office-js-docs-reference/tree/master/docs/docs-ref-autogen) folder, please [submit an issue](https://github.com/OfficeDev/office-js-docs-reference/issues) in this repository.</span></span> <span data-ttu-id="80b00-115">Подробнее о том, как работает инструментарий в этом репозитории, можно узнать [здесь.](https://github.com/OfficeDev/office-js-docs-reference/blob/master/DocumentationToolingNotes.md)</span><span class="sxs-lookup"><span data-stu-id="80b00-115">You can read more about how the tooling in this repository [here](https://github.com/OfficeDev/office-js-docs-reference/blob/master/DocumentationToolingNotes.md).</span></span>

* <span data-ttu-id="80b00-116">Чтобы узнать о вашем опыте программирования, о том, что вы хотели бы видеть в будущих версиях, примерах кода и так далее, введите свои предложения и идеи в [Microsoft 365](https://docs.microsoft.com/answers/products/m365)на Q&A .</span><span class="sxs-lookup"><span data-stu-id="80b00-116">To let us know about your programming experience, what you would like to see in future versions, code samples, and so on, enter your suggestions and ideas under [Microsoft 365 on Q&A](https://docs.microsoft.com/answers/products/m365).</span></span>

## <a name="join-the-microsoft-365-developer-program"></a><span data-ttu-id="80b00-117">Присоединяйтесь к программе разработчиков Microsoft 365</span><span class="sxs-lookup"><span data-stu-id="80b00-117">Join the Microsoft 365 Developer Program</span></span>

<span data-ttu-id="80b00-118">Получите бесплатную песочницу, инструменты и другие ресурсы, необходимые для создания решений для платформы Microsoft 365.</span><span class="sxs-lookup"><span data-stu-id="80b00-118">Get a free sandbox, tools, and other resources you need to build solutions for the Microsoft 365 platform.</span></span>

* <span data-ttu-id="80b00-119">[Бесплатная песочница разработчика](https://developer.microsoft.com/microsoft-365/dev-program#Subscription) Получите бесплатную, возобновляемую 90-дневную подписку на разработчика Microsoft 365 E5.</span><span class="sxs-lookup"><span data-stu-id="80b00-119">[Free developer sandbox](https://developer.microsoft.com/microsoft-365/dev-program#Subscription) Get a free, renewable 90-day Microsoft 365 E5 developer subscription.</span></span>
* <span data-ttu-id="80b00-120">[Примеры пакетов данных](https://developer.microsoft.com/microsoft-365/dev-program#Sample) Автоматически настраивайте песочницу, устанавливая пользовательские данные и контент для создания решений.</span><span class="sxs-lookup"><span data-stu-id="80b00-120">[Sample data packs](https://developer.microsoft.com/microsoft-365/dev-program#Sample) Automatically configure your sandbox by installing user data and content to help you build your solutions.</span></span>
* <span data-ttu-id="80b00-121">[Доступ к экспертам](https://developer.microsoft.com/microsoft-365/dev-program#Experts) Доступ к событиям сообщества, чтобы узнать у экспертов Microsoft 365.</span><span class="sxs-lookup"><span data-stu-id="80b00-121">[Access to experts](https://developer.microsoft.com/microsoft-365/dev-program#Experts) Access community events to learn from Microsoft 365 experts.</span></span>
* <span data-ttu-id="80b00-122">[Персонализированные рекомендации](https://developer.microsoft.com/microsoft-365/dev-program#Recommendations) Быстро найдите ресурсы разработчика из персонализированной панели мониторинга.</span><span class="sxs-lookup"><span data-stu-id="80b00-122">[Personalized recommendations](https://developer.microsoft.com/microsoft-365/dev-program#Recommendations) Find developer resources quickly from your personalized dashboard.</span></span>


## <a name="microsoft-open-source-code-of-conduct"></a><span data-ttu-id="80b00-123">Правила поведения Майкрософт, касающиеся обращения с открытым кодом</span><span class="sxs-lookup"><span data-stu-id="80b00-123">Microsoft Open Source Code of Conduct</span></span>

<span data-ttu-id="80b00-124">Этот проект соответствует [правилам поведения Майкрософт, касающимся обращения с открытым кодом](https://opensource.microsoft.com/codeofconduct/).</span><span class="sxs-lookup"><span data-stu-id="80b00-124">This project has adopted the [Microsoft Open Source Code of Conduct](https://opensource.microsoft.com/codeofconduct/).</span></span>
<span data-ttu-id="80b00-125">Дополнительные сведения см. в [разделе Вопросы](https://opensource.microsoft.com/codeofconduct/faq/)кодекса поведения или свяжитесь с opencode@microsoft.com [дополнительными](mailto:opencode@microsoft.com) вопросами или комментариями.</span><span class="sxs-lookup"><span data-stu-id="80b00-125">For more information, see the [Code of Conduct FAQ](https://opensource.microsoft.com/codeofconduct/faq/), or contact [opencode@microsoft.com](mailto:opencode@microsoft.com) with any additional questions or comments.</span></span>
