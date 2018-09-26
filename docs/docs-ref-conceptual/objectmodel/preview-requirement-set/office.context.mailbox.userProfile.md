
# <a name="userprofile"></a>userProfile

### [Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md). userProfile

##### <a name="requirements"></a>Требования

|Requirement| Значение|
|---|---|
|[Версия минимального набора требований к почтовому ящику](/javascript/office/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[Минимальный уровень разрешений](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem|
|[Применимый режим Outlook](https://docs.microsoft.com/outlook/add-ins/#extension-points)| Создание или чтение|

##### <a name="members-and-methods"></a>Элементы и методы

| Элемент | Тип |
|--------|------|
| [accountType](#accounttype-string) | Member |
| [displayName](#displayname-string) | Элемент |
| [emailAddress](#emailaddress-string) | Элемент |
| [timeZone](#timezone-string) | Элемент |

### <a name="members"></a>Members

####  <a name="accounttype-string"></a>accountType: String

> [!NOTE]
> Этот член является в настоящее время только поддерживаемые в Outlook 2016 или более поздней версии для Mac (построение 16.9.1212 или более поздней версии).

Получает тип учетной записи пользователя, связанного с почтовым ящиком. В следующей таблице перечислены возможные значения.

| Значение | Описание |
|-------|-------------|
| `enterprise` | Почтовый ящик относится локального сервера Exchange. |
| `gmail` | Почтовый ящик связан с учетной записью Gmail. |
| `office365` | Почтовый ящик связан с Office 365 работы или школе учетной записи. |
| `outlookCom` | Почтовый ящик связан с учетной записью личных Outlook.com. |

##### <a name="type"></a>Тип:

*   String

##### <a name="requirements"></a>Требования

|Requirement| Значение|
|---|---|
|[Версия минимального набора требований к почтовому ящику](/javascript/office/requirement-sets/outlook-api-requirement-sets)| 1.6 |
|[Минимальный уровень разрешений](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem|
|[Применимый режим Outlook](https://docs.microsoft.com/outlook/add-ins/#extension-points)| Создание или чтение|

##### <a name="example"></a>Пример

```
console.log(Office.context.mailbox.userProfile.accountType);
```

####  <a name="displayname-string"></a>displayName :String

Получает отображаемое имя пользователя.

##### <a name="type"></a>Тип:

*   String

##### <a name="requirements"></a>Требования

|Requirement| Значение|
|---|---|
|[Версия минимального набора требований к почтовому ящику](/javascript/office/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[Минимальный уровень разрешений](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem|
|[Применимый режим Outlook](https://docs.microsoft.com/outlook/add-ins/#extension-points)| Создание или чтение|

##### <a name="example"></a>Пример

```
// Example: Allie Bellew
console.log(Office.context.mailbox.userProfile.displayName);
```

####  <a name="emailaddress-string"></a>emailAddress :String

Получает адрес электронной почты SMTP пользователя.

##### <a name="type"></a>Тип:

*   String

##### <a name="requirements"></a>Требования

|Requirement| Значение|
|---|---|
|[Версия минимального набора требований к почтовому ящику](/javascript/office/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[Минимальный уровень разрешений](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem|
|[Применимый режим Outlook](https://docs.microsoft.com/outlook/add-ins/#extension-points)| Создание или чтение|

##### <a name="example"></a>Пример

```
// Example: allieb@contoso.com
console.log(Office.context.mailbox.userProfile.emailAddress);
```

####  <a name="timezone-string"></a>timeZone :String

Получает часовой пояс пользователя по умолчанию.

##### <a name="type"></a>Тип:

*   String

##### <a name="requirements"></a>Требования

|Requirement| Значение|
|---|---|
|[Версия минимального набора требований к почтовому ящику](/javascript/office/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[Минимальный уровень разрешений](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem|
|[Применимый режим Outlook](https://docs.microsoft.com/outlook/add-ins/#extension-points)| Создание или чтение|

##### <a name="example"></a>Пример

```
// Example: Pacific Standard Time
console.log(Office.context.mailbox.userProfile.timeZone);
```