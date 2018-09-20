# <a name="identity-api-requirement-sets"></a>Наборы обязательных элементов API удостоверений

Наборы обязательных элементов — это именованные группы элементов API. Надстройки Office использовать наборов требований, указанный в манифесте или выполняется проверка среды выполнения для определения поддержки API, которые требуется добавить в приложение Office. Дополнительные сведения см в [различных версиях Office и требования наборов](https://docs.microsoft.com/office/dev/add-ins/develop/office-versions-and-requirement-sets).

Запуск надстроек Office в разных версиях Office. В следующей таблице перечислены наборы требований Identity API ведущие приложения Office, которые поддерживают числа, set и построения или версии требования для приложения Office.

|  Набор обязательных элементов  | Office 2013 для Windows | Office 365 для Windows   |  Office 365 для iPad  |  Office 365 для Mac  | Office Online  | SharePoint Online | OneDrive.com |Outlook.com и Exchange Online|
|:-----|-----|:-----|:-----|:-----|:-----|:-----|:-----|:-----|
| IdentityAPI 1.1  | Н/Д | Предварительная версия **& #42;** | Ожидается в скором времени | Предварительная версия **& #42;**| Доступно | Доступно| Ожидается в скором времени | Ожидается в скором времени |

> **& #42;** На этапе предварительного просмотра API удостоверений поддерживается на Windows 2016 и Mac только для пользователей в программе сотрудниками, используя быстрый режим. Чтобы присоединиться к программе сотрудники компании, обратитесь к [быть где Office](https://products.office.com/office-insider?tab=tab-1). Чтобы переключиться на быстро увидеть [Изнутри Fast](https://answers.microsoft.com/en-us/msoffice/forum/msoffice_officeinsider-mso_win10-msoinsider_reg/its-here-office-insider-fast-for-office-2016-on/dbe8e7bb-9523-44a4-948b-9436fedfd961).

Статьи и разделы с дополнительными сведениями о версиях, номерах сборок и Office Online Server:

- [Номера версий и сборок выпусков из канала обновления для клиентов Office 365](https://support.office.com/article/version-and-build-numbers-of-update-channel-releases-ae942449-1fca-4484-898b-a933ea23def7);
- [Какая у меня версия Office](https://support.office.com/article/What-version-of-Office-am-I-using-932788b8-a3ce-44bf-bb09-e334518b8b19);
- [Где можно найти номера версии и сборки клиентского приложения Office 365](https://support.office.com/article/version-and-build-numbers-of-update-channel-releases-ae942449-1fca-4484-898b-a933ea23def7);
- 
  [Обзор Office Online Server](https://docs.microsoft.com/officeonlineserver/office-online-server-overview).

## <a name="office-common-api-requirement-sets"></a>Стандартные наборы обязательных элементов API для Office

Сведения о наборах обязательных элементов общего API для Office см. в [этой статье](office-add-in-requirement-sets.md).

## <a name="identityapi-11"></a>IdentityAPI 1.1 

IdentityAPI 1.1 для единого входа — это первая версия API. Для получения дополнительных сведений об API увидеть `getAccessTokenAsync` метод в справочном разделе [Office.Auth](/javascript/api/office/office.auth) .

## <a name="see-also"></a>См. также

- [Версии Office и наборы обязательных элементов](https://docs.microsoft.com/office/dev/add-ins/develop/office-versions-and-requirement-sets)
- [Указание ведущих приложений Office и обязательных элементов API](https://docs.microsoft.com/office/dev/add-ins/develop/specify-office-hosts-and-api-requirements)
- [XML-манифест надстроек Office](https://docs.microsoft.com/office/dev/add-ins/develop/add-in-manifests)
