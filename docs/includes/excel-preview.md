| Класс | Поля | Описание |
|:---|:---|:---|
|[Объект AllowEditRange](/javascript/api/excel/excel.alloweditrange)|[address](/javascript/api/excel/excel.alloweditrange#excel-excel-alloweditrange-address-member)|Указывает диапазон, связанный с объектом.|
||[delete()](/javascript/api/excel/excel.alloweditrange#excel-excel-alloweditrange-delete-member(1))|Удаляет этот объект из `AllowEditRangeCollection`.|
||[isPasswordProtected](/javascript/api/excel/excel.alloweditrange#excel-excel-alloweditrange-ispasswordprotected-member)|Указывает, защищен ли `AllowEditRange` пароль.|
||[pauseProtection(password?: string)](/javascript/api/excel/excel.alloweditrange#excel-excel-alloweditrange-pauseprotection-member(1))|Приостановка защиты таблиц для данного `AllowEditRange` объекта для пользователя в заданном сеансе.|
||[setPassword(password?: string)](/javascript/api/excel/excel.alloweditrange#excel-excel-alloweditrange-setpassword-member(1))|Изменяет пароль, связанный с `AllowEditRange`.|
||[заголовок](/javascript/api/excel/excel.alloweditrange#excel-excel-alloweditrange-title-member)|Указывает название объекта.|
|[AllowEditRangeCollection](/javascript/api/excel/excel.alloweditrangecollection)|[add(title: string, rangeAddress: string, options?: Excel. AllowEditRangeOptions)](/javascript/api/excel/excel.alloweditrangecollection#excel-excel-alloweditrangecollection-add-member(1))|Добавляет объект `AllowEditRange` в коллекцию.|
||[getCount()](/javascript/api/excel/excel.alloweditrangecollection#excel-excel-alloweditrangecollection-getcount-member(1))|Возвращает количество объектов `AllowEditRange` в коллекции.|
||[getItem(key: string)](/javascript/api/excel/excel.alloweditrangecollection#excel-excel-alloweditrangecollection-getitem-member(1))|Получает объект `AllowEditRange` по его названию.|
||[getItemAt(index: number)](/javascript/api/excel/excel.alloweditrangecollection#excel-excel-alloweditrangecollection-getitemat-member(1))|Возвращает объект `AllowEditRange` по индексу в коллекции.|
||[getItemOrNullObject(key: string)](/javascript/api/excel/excel.alloweditrangecollection#excel-excel-alloweditrangecollection-getitemornullobject-member(1))|Получает объект `AllowEditRange` по его названию.|
||[items](/javascript/api/excel/excel.alloweditrangecollection#excel-excel-alloweditrangecollection-items-member)|Получает загруженные дочерние элементы в этой коллекции.|
||[pauseProtection (пароль: строка)](/javascript/api/excel/excel.alloweditrangecollection#excel-excel-alloweditrangecollection-pauseprotection-member(1))|Приостановка защиты от таблиц для всех `AllowEditRange` объектов в коллекции с заданным паролем для пользователя в заданном сеансе.|
|[AllowEditRangeOptions](/javascript/api/excel/excel.alloweditrangeoptions)|[password](/javascript/api/excel/excel.alloweditrangeoptions#excel-excel-alloweditrangeoptions-password-member)|Пароль, связанный с `AllowEditRange`.|
|[ArrayCellValue](/javascript/api/excel/excel.arraycellvalue)|[basicType](/javascript/api/excel/excel.arraycellvalue#excel-excel-arraycellvalue-basictype-member)|Представляет значение, которое будет возвращено ячейкой `Range.valueTypes` с этим значением.|
||[basicValue](/javascript/api/excel/excel.arraycellvalue#excel-excel-arraycellvalue-basicvalue-member)|Представляет значение, которое будет возвращено ячейкой `Range.values` с этим значением.|
||[элементы](/javascript/api/excel/excel.arraycellvalue#excel-excel-arraycellvalue-elements-member)|Представляет элементы массива.|
||[referencedValues](/javascript/api/excel/excel.arraycellvalue#excel-excel-arraycellvalue-referencedvalues-member)|Представляет значения ячейки, которые ссылались в пределах `ArrayCellValue.elements`.|
||[тип](/javascript/api/excel/excel.arraycellvalue#excel-excel-arraycellvalue-type-member)|Представляет тип этого значения ячейки.|
|[BlockedErrorCellValue](/javascript/api/excel/excel.blockederrorcellvalue)|[basicType](/javascript/api/excel/excel.blockederrorcellvalue#excel-excel-blockederrorcellvalue-basictype-member)|Представляет значение, которое будет возвращено ячейкой `Range.valueTypes` с этим значением.|
||[basicValue](/javascript/api/excel/excel.blockederrorcellvalue#excel-excel-blockederrorcellvalue-basicvalue-member)|Представляет значение, которое будет возвращено ячейкой `Range.values` с этим значением.|
||[errorSubType](/javascript/api/excel/excel.blockederrorcellvalue#excel-excel-blockederrorcellvalue-errorsubtype-member)|Представляет тип `BlockedErrorCellValue`.|
||[errorType](/javascript/api/excel/excel.blockederrorcellvalue#excel-excel-blockederrorcellvalue-errortype-member)|Представляет тип `ErrorCellValue`.|
||[тип](/javascript/api/excel/excel.blockederrorcellvalue#excel-excel-blockederrorcellvalue-type-member)|Представляет тип этого значения ячейки.|
|[BooleanCellValue](/javascript/api/excel/excel.booleancellvalue)|[basicType](/javascript/api/excel/excel.booleancellvalue#excel-excel-booleancellvalue-basictype-member)|Представляет значение, которое будет возвращено ячейкой `Range.valueTypes` с этим значением.|
||[basicValue](/javascript/api/excel/excel.booleancellvalue#excel-excel-booleancellvalue-basicvalue-member)|Представляет значение, которое будет возвращено ячейкой `Range.values` с этим значением.|
||[тип](/javascript/api/excel/excel.booleancellvalue#excel-excel-booleancellvalue-type-member)|Представляет тип этого значения ячейки.|
|[BusyErrorCellValue](/javascript/api/excel/excel.busyerrorcellvalue)|[basicType](/javascript/api/excel/excel.busyerrorcellvalue#excel-excel-busyerrorcellvalue-basictype-member)|Представляет значение, которое будет возвращено ячейкой `Range.valueTypes` с этим значением.|
||[basicValue](/javascript/api/excel/excel.busyerrorcellvalue#excel-excel-busyerrorcellvalue-basicvalue-member)|Представляет значение, которое будет возвращено ячейкой `Range.values` с этим значением.|
||[errorSubType](/javascript/api/excel/excel.busyerrorcellvalue#excel-excel-busyerrorcellvalue-errorsubtype-member)|Представляет тип `BusyErrorCellValue`.|
||[errorType](/javascript/api/excel/excel.busyerrorcellvalue#excel-excel-busyerrorcellvalue-errortype-member)|Представляет тип `ErrorCellValue`.|
||[тип](/javascript/api/excel/excel.busyerrorcellvalue#excel-excel-busyerrorcellvalue-type-member)|Представляет тип этого значения ячейки.|
|[CalcErrorCellValue](/javascript/api/excel/excel.calcerrorcellvalue)|[basicType](/javascript/api/excel/excel.calcerrorcellvalue#excel-excel-calcerrorcellvalue-basictype-member)|Представляет значение, которое будет возвращено ячейкой `Range.valueTypes` с этим значением.|
||[basicValue](/javascript/api/excel/excel.calcerrorcellvalue#excel-excel-calcerrorcellvalue-basicvalue-member)|Представляет значение, которое будет возвращено ячейкой `Range.values` с этим значением.|
||[errorSubType](/javascript/api/excel/excel.calcerrorcellvalue#excel-excel-calcerrorcellvalue-errorsubtype-member)|Представляет тип `CalcErrorCellValue`.|
||[errorType](/javascript/api/excel/excel.calcerrorcellvalue#excel-excel-calcerrorcellvalue-errortype-member)|Представляет тип `ErrorCellValue`.|
||[тип](/javascript/api/excel/excel.calcerrorcellvalue#excel-excel-calcerrorcellvalue-type-member)|Представляет тип этого значения ячейки.|
|[CardLayoutListSection](/javascript/api/excel/excel.cardlayoutlistsection)|[макет](/javascript/api/excel/excel.cardlayoutlistsection#excel-excel-cardlayoutlistsection-layout-member)|Представляет тип макета для этого раздела.|
|[CardLayoutPropertyReference](/javascript/api/excel/excel.cardlayoutpropertyreference)|[property](/javascript/api/excel/excel.cardlayoutpropertyreference#excel-excel-cardlayoutpropertyreference-property-member)|Имя свойства, на которое ссылается макет карты.|
|[CardLayoutSectionStandardProperties](/javascript/api/excel/excel.cardlayoutsectionstandardproperties)|[обвалился](/javascript/api/excel/excel.cardlayoutsectionstandardproperties#excel-excel-cardlayoutsectionstandardproperties-collapsed-member)|Представляет, был ли этот раздел карты первоначально свернут.|
||[collapsible](/javascript/api/excel/excel.cardlayoutsectionstandardproperties#excel-excel-cardlayoutsectionstandardproperties-collapsible-member)|Представляет, является ли этот раздел карты коллапсируемым.|
||[properties](/javascript/api/excel/excel.cardlayoutsectionstandardproperties#excel-excel-cardlayoutsectionstandardproperties-properties-member)|Представляет имена свойств в этом разделе.|
||[заголовок](/javascript/api/excel/excel.cardlayoutsectionstandardproperties#excel-excel-cardlayoutsectionstandardproperties-title-member)|Представляет название этого раздела карты.|
|[CardLayoutStandardProperties](/javascript/api/excel/excel.cardlayoutstandardproperties)|[mainImage](/javascript/api/excel/excel.cardlayoutstandardproperties#excel-excel-cardlayoutstandardproperties-mainimage-member)|Указывает свойство, которое будет использоваться в качестве основного изображения карты.|
||[sections](/javascript/api/excel/excel.cardlayoutstandardproperties#excel-excel-cardlayoutstandardproperties-sections-member)|Представляет разделы карты.|
||[subTitle](/javascript/api/excel/excel.cardlayoutstandardproperties#excel-excel-cardlayoutstandardproperties-subtitle-member)|Представляет спецификацию, в которой свойство содержит подзаголовок карты.|
||[заголовок](/javascript/api/excel/excel.cardlayoutstandardproperties#excel-excel-cardlayoutstandardproperties-title-member)|Представляет название карты или спецификацию, в которой свойство содержит название карты.|
|[CardLayoutTableSection](/javascript/api/excel/excel.cardlayouttablesection)|[макет](/javascript/api/excel/excel.cardlayouttablesection#excel-excel-cardlayouttablesection-layout-member)|Представляет тип макета для этого раздела.|
|[CellValueAttributionAttributes](/javascript/api/excel/excel.cellvalueattributionattributes)|[licenseAddress](/javascript/api/excel/excel.cellvalueattributionattributes#excel-excel-cellvalueattributionattributes-licenseaddress-member)|Представляет URL-адрес лицензии или источника, описывая, как можно использовать это свойство.|
||[licenseText](/javascript/api/excel/excel.cellvalueattributionattributes#excel-excel-cellvalueattributionattributes-licensetext-member)|Представляет имя лицензии, управляющей этим свойством.|
||[sourceAddress](/javascript/api/excel/excel.cellvalueattributionattributes#excel-excel-cellvalueattributionattributes-sourceaddress-member)|Представляет URL-адрес источника `CellValue`.|
||[sourceText](/javascript/api/excel/excel.cellvalueattributionattributes#excel-excel-cellvalueattributionattributes-sourcetext-member)|Представляет имя источника `CellValue`.|
|[CellValueExtraProperties](/javascript/api/excel/excel.cellvalueextraproperties)|[writable](/javascript/api/excel/excel.cellvalueextraproperties#excel-excel-cellvalueextraproperties-writable-member)|Представляет, будет ли `CellValue` это использоваться для переописывание ячейки.|
||[writableNote](/javascript/api/excel/excel.cellvalueextraproperties#excel-excel-cellvalueextraproperties-writablenote-member)|Представляет объяснение того, почему `CellValue.writable` указывается как false.|
|[CellValuePropertyMetadata](/javascript/api/excel/excel.cellvaluepropertymetadata)|[атрибуция](/javascript/api/excel/excel.cellvaluepropertymetadata#excel-excel-cellvaluepropertymetadata-attribution-member)|Представляет сведения о присвоении для описания исходных и лицензионных требований для использования этого свойства.|
||[excludeFrom](/javascript/api/excel/excel.cellvaluepropertymetadata#excel-excel-cellvaluepropertymetadata-excludefrom-member)|Представляет, какие функции этого свойства исключены из.|
||[sublabel](/javascript/api/excel/excel.cellvaluepropertymetadata#excel-excel-cellvaluepropertymetadata-sublabel-member)|Представляет подлабель для этого свойства, показанного в представлении карты.|
|[CellValuePropertyMetadataExclusions](/javascript/api/excel/excel.cellvaluepropertymetadataexclusions)|[autoComplete](/javascript/api/excel/excel.cellvaluepropertymetadataexclusions#excel-excel-cellvaluepropertymetadataexclusions-autocomplete-member)|True означает, что свойство исключено из свойств, показанных автоматически завершенным.|
||[calcCompare](/javascript/api/excel/excel.cellvaluepropertymetadataexclusions#excel-excel-cellvaluepropertymetadataexclusions-calccompare-member)|True представляет, что свойство исключено из свойств, используемых для сравнения значений ячейки во время пересчета.|
||[cardView](/javascript/api/excel/excel.cellvaluepropertymetadataexclusions#excel-excel-cellvaluepropertymetadataexclusions-cardview-member)|True означает, что свойство исключено из свойств, показанных в представлении карты.|
||[dotNotation](/javascript/api/excel/excel.cellvaluepropertymetadataexclusions#excel-excel-cellvaluepropertymetadataexclusions-dotnotation-member)|True означает, что свойство исключено из свойств, к которым можно получить доступ с помощью функции FIELDVALUE.|
|[CellValueProviderAttributes](/javascript/api/excel/excel.cellvalueproviderattributes)|[description](/javascript/api/excel/excel.cellvalueproviderattributes#excel-excel-cellvalueproviderattributes-description-member)|Представляет свойство описания поставщика, используемое в представлении карты, если не указан логотип.|
||[logoSourceAddress](/javascript/api/excel/excel.cellvalueproviderattributes#excel-excel-cellvalueproviderattributes-logosourceaddress-member)|Представляет URL-адрес, используемый для загрузки изображения, которое будет использоваться в качестве логотипа в представлении карты.|
||[logoTargetAddress](/javascript/api/excel/excel.cellvalueproviderattributes#excel-excel-cellvalueproviderattributes-logotargetaddress-member)|Представляет URL-адрес, который является объектом навигации, если пользователь щелкает элементом логотипа в представлении карты.|
|[ChartFill](/javascript/api/excel/excel.chartfill)|[solidColor](/javascript/api/excel/excel.chartfill#excel-excel-chartfill-solidcolor-member)|Задает и получает форматирование однообразной цветовой заливки элемента диаграммы.|
|[ChartSeries](/javascript/api/excel/excel.chartseries)|[getDimensionDataSourceString(размер: Excel. ChartSeriesDimension)](/javascript/api/excel/excel.chartseries#excel-excel-chartseries-getdimensiondatasourcestring-member(1))|Получает строковую репрезентативность источника данных серии диаграмм. Представление строки может быть информацией, например адресом ячейки.|
||[getDimensionDataSourceType(размер: Excel. ChartSeriesDimension)](/javascript/api/excel/excel.chartseries#excel-excel-chartseries-getdimensiondatasourcetype-member(1))|Получает тип источника данных из серии диаграмм.|
|[Comment](/javascript/api/excel/excel.comment)|[assignTask(assignee: Identity)](/javascript/api/excel/excel.comment#excel-excel-comment-assigntask-member(1))|Назначает задачу, прикрепленную к комментарию, для данного пользователя в качестве ассимилята.|
||[getTask()](/javascript/api/excel/excel.comment#excel-excel-comment-gettask-member(1))|Получает задачу, связанную с этим комментарием.|
||[getTaskOrNullObject()](/javascript/api/excel/excel.comment#excel-excel-comment-gettaskornullobject-member(1))|Получает задачу, связанную с этим комментарием.|
|[CommentReply](/javascript/api/excel/excel.commentreply)|[assignTask(assignee: Identity)](/javascript/api/excel/excel.commentreply#excel-excel-commentreply-assigntask-member(1))|Назначает задачу, прикрепленную к комментарию, для данного пользователя в качестве единственного назначаемой.|
||[getTask()](/javascript/api/excel/excel.commentreply#excel-excel-commentreply-gettask-member(1))|Получает задачу, связанную с потоком ответа на этот комментарий.|
||[getTaskOrNullObject()](/javascript/api/excel/excel.commentreply#excel-excel-commentreply-gettaskornullobject-member(1))|Получает задачу, связанную с потоком ответа на этот комментарий.|
|[ConnectErrorCellValue](/javascript/api/excel/excel.connecterrorcellvalue)|[basicType](/javascript/api/excel/excel.connecterrorcellvalue#excel-excel-connecterrorcellvalue-basictype-member)|Представляет значение, которое будет возвращено ячейкой `Range.valueTypes` с этим значением.|
||[basicValue](/javascript/api/excel/excel.connecterrorcellvalue#excel-excel-connecterrorcellvalue-basicvalue-member)|Представляет значение, которое будет возвращено ячейкой `Range.values` с этим значением.|
||[errorSubType](/javascript/api/excel/excel.connecterrorcellvalue#excel-excel-connecterrorcellvalue-errorsubtype-member)|Представляет тип `ConnectErrorCellValue`.|
||[errorType](/javascript/api/excel/excel.connecterrorcellvalue#excel-excel-connecterrorcellvalue-errortype-member)|Представляет тип `ErrorCellValue`.|
||[тип](/javascript/api/excel/excel.connecterrorcellvalue#excel-excel-connecterrorcellvalue-type-member)|Представляет тип этого значения ячейки.|
|[Div0ErrorCellValue](/javascript/api/excel/excel.div0errorcellvalue)|[basicType](/javascript/api/excel/excel.div0errorcellvalue#excel-excel-div0errorcellvalue-basictype-member)|Представляет значение, которое будет возвращено ячейкой `Range.valueTypes` с этим значением.|
||[basicValue](/javascript/api/excel/excel.div0errorcellvalue#excel-excel-div0errorcellvalue-basicvalue-member)|Представляет значение, которое будет возвращено ячейкой `Range.values` с этим значением.|
||[errorType](/javascript/api/excel/excel.div0errorcellvalue#excel-excel-div0errorcellvalue-errortype-member)|Представляет тип `ErrorCellValue`.|
||[тип](/javascript/api/excel/excel.div0errorcellvalue#excel-excel-div0errorcellvalue-type-member)|Представляет тип этого значения ячейки.|
|[DocumentTask](/javascript/api/excel/excel.documenttask)|[назначение](/javascript/api/excel/excel.documenttask#excel-excel-documenttask-assignees-member)|Возвращает коллекцию назначений задачи.|
||[изменения](/javascript/api/excel/excel.documenttask#excel-excel-documenttask-changes-member)|Получает записи изменений задачи.|
||[comment](/javascript/api/excel/excel.documenttask#excel-excel-documenttask-comment-member)|Получает комментарий, связанный с задачей.|
||[completedBy](/javascript/api/excel/excel.documenttask#excel-excel-documenttask-completedby-member)|Получает последнего пользователя, который выполнил задачу.|
||[completedDateTime](/javascript/api/excel/excel.documenttask#excel-excel-documenttask-completeddatetime-member)|Получает дату и время завершения задачи.|
||[createdBy](/javascript/api/excel/excel.documenttask#excel-excel-documenttask-createdby-member)|Получает пользователя, создавшего задачу.|
||[createdDateTime](/javascript/api/excel/excel.documenttask#excel-excel-documenttask-createddatetime-member)|Получает дату и время создания задачи.|
||[id](/javascript/api/excel/excel.documenttask#excel-excel-documenttask-id-member)|Получает ID задачи.|
||[percentComplete](/javascript/api/excel/excel.documenttask#excel-excel-documenttask-percentcomplete-member)|Указывает процент выполнения задачи.|
||[приоритет](/javascript/api/excel/excel.documenttask#excel-excel-documenttask-priority-member)|Указывает приоритет задачи.|
||[setStartAndDueDateTime (startDateTime: Date, dueDateTime: Date)](/javascript/api/excel/excel.documenttask#excel-excel-documenttask-setstartandduedatetime-member(1))|Изменяет даты начала и срока действия задачи.|
||[startAndDueDateTime](/javascript/api/excel/excel.documenttask#excel-excel-documenttask-startandduedatetime-member)|Получает или задает дату и время, когда должна начаться и должна быть поставлена задача.|
||[заголовок](/javascript/api/excel/excel.documenttask#excel-excel-documenttask-title-member)|Указывает название задачи.|
|[DocumentTaskChange](/javascript/api/excel/excel.documenttaskchange)|[assignee](/javascript/api/excel/excel.documenttaskchange#excel-excel-documenttaskchange-assignee-member)|Представляет пользователя, назначенного `assign` для задачи для типа записи изменений, или пользователя, не назначенного из задачи для типа записи `unassign` изменений.|
||[changedBy](/javascript/api/excel/excel.documenttaskchange#excel-excel-documenttaskchange-changedby-member)|Представляет пользователя, создавшего или измениввшего задачу.|
||[commentId](/javascript/api/excel/excel.documenttaskchange#excel-excel-documenttaskchange-commentid-member)|Представляет ID того или `Comment` иного изменения `CommentReply` задачи.|
||[createdDateTime](/javascript/api/excel/excel.documenttaskchange#excel-excel-documenttaskchange-createddatetime-member)|Представляет дату создания и время записи изменения задачи.|
||[dueDateTime](/javascript/api/excel/excel.documenttaskchange#excel-excel-documenttaskchange-duedatetime-member)|Представляет дату и время задачи в часовом поясе UTC.|
||[id](/javascript/api/excel/excel.documenttaskchange#excel-excel-documenttaskchange-id-member)|ID для записи изменения задачи.|
||[percentComplete](/javascript/api/excel/excel.documenttaskchange#excel-excel-documenttaskchange-percentcomplete-member)|Представляет процент выполнения задачи.|
||[приоритет](/javascript/api/excel/excel.documenttaskchange#excel-excel-documenttaskchange-priority-member)|Представляет приоритет задачи.|
||[startDateTime](/javascript/api/excel/excel.documenttaskchange#excel-excel-documenttaskchange-startdatetime-member)|Представляет дату и время начала задачи в часовом поясе UTC.|
||[заголовок](/javascript/api/excel/excel.documenttaskchange#excel-excel-documenttaskchange-title-member)|Представляет название задачи.|
||[тип](/javascript/api/excel/excel.documenttaskchange#excel-excel-documenttaskchange-type-member)|Представляет тип действия записи изменения задачи.|
||[undoHistoryId](/javascript/api/excel/excel.documenttaskchange#excel-excel-documenttaskchange-undohistoryid-member)|Представляет свойство `DocumentTaskChange.id` , которое было отменено для типа записи `undo` изменений.|
|[DocumentTaskChangeCollection](/javascript/api/excel/excel.documenttaskchangecollection)|[getCount()](/javascript/api/excel/excel.documenttaskchangecollection#excel-excel-documenttaskchangecollection-getcount-member(1))|Получает количество записей изменений в коллекции для задачи.|
||[getItemAt(index: number)](/javascript/api/excel/excel.documenttaskchangecollection#excel-excel-documenttaskchangecollection-getitemat-member(1))|Получает запись изменения задачи с помощью индекса в коллекции.|
||[items](/javascript/api/excel/excel.documenttaskchangecollection#excel-excel-documenttaskchangecollection-items-member)|Получает загруженные дочерние элементы в этой коллекции.|
|[DocumentTaskCollection](/javascript/api/excel/excel.documenttaskcollection)|[getCount()](/javascript/api/excel/excel.documenttaskcollection#excel-excel-documenttaskcollection-getcount-member(1))|Получает количество задач в коллекции.|
||[getItem(key: string)](/javascript/api/excel/excel.documenttaskcollection#excel-excel-documenttaskcollection-getitem-member(1))|Получает задачу с помощью своего ID.|
||[getItemAt(index: number)](/javascript/api/excel/excel.documenttaskcollection#excel-excel-documenttaskcollection-getitemat-member(1))|Получает задачу по индексу в коллекции.|
||[getItemOrNullObject(key: string)](/javascript/api/excel/excel.documenttaskcollection#excel-excel-documenttaskcollection-getitemornullobject-member(1))|Получает задачу с помощью своего ID.|
||[items](/javascript/api/excel/excel.documenttaskcollection#excel-excel-documenttaskcollection-items-member)|Получает загруженные дочерние элементы в этой коллекции.|
|[DocumentTaskSchedule](/javascript/api/excel/excel.documenttaskschedule)|[dueDateTime](/javascript/api/excel/excel.documenttaskschedule#excel-excel-documenttaskschedule-duedatetime-member)|Получает дату и время, когда должна быть поставлена задача.|
||[startDateTime](/javascript/api/excel/excel.documenttaskschedule#excel-excel-documenttaskschedule-startdatetime-member)|Получает дату и время, которые должна начаться задача.|
|[DoubleCellValue](/javascript/api/excel/excel.doublecellvalue)|[basicType](/javascript/api/excel/excel.doublecellvalue#excel-excel-doublecellvalue-basictype-member)|Представляет значение, которое будет возвращено ячейкой `Range.valueTypes` с этим значением.|
||[basicValue](/javascript/api/excel/excel.doublecellvalue#excel-excel-doublecellvalue-basicvalue-member)|Представляет значение, которое будет возвращено ячейкой `Range.values` с этим значением.|
||[тип](/javascript/api/excel/excel.doublecellvalue#excel-excel-doublecellvalue-type-member)|Представляет тип этого значения ячейки.|
|[EmptyCellValue](/javascript/api/excel/excel.emptycellvalue)|[basicType](/javascript/api/excel/excel.emptycellvalue#excel-excel-emptycellvalue-basictype-member)|Представляет значение, которое будет возвращено ячейкой `Range.valueTypes` с этим значением.|
||[basicValue](/javascript/api/excel/excel.emptycellvalue#excel-excel-emptycellvalue-basicvalue-member)|Представляет значение, которое будет возвращено ячейкой `Range.values` с этим значением.|
||[тип](/javascript/api/excel/excel.emptycellvalue#excel-excel-emptycellvalue-type-member)|Представляет тип этого значения ячейки.|
|[EntityCardLayout](/javascript/api/excel/excel.entitycardlayout)|[макет](/javascript/api/excel/excel.entitycardlayout#excel-excel-entitycardlayout-layout-member)|Представляет тип этого макета.|
|[EntityCellValue](/javascript/api/excel/excel.entitycellvalue)|[basicType](/javascript/api/excel/excel.entitycellvalue#excel-excel-entitycellvalue-basictype-member)|Представляет значение, которое будет возвращено ячейкой `Range.valueTypes` с этим значением.|
||[basicValue](/javascript/api/excel/excel.entitycellvalue#excel-excel-entitycellvalue-basicvalue-member)|Представляет значение, которое будет возвращено ячейкой `Range.values` с этим значением.|
||[макеты](/javascript/api/excel/excel.entitycellvalue#excel-excel-entitycellvalue-layouts-member)|Представляет сведения о макете для представлений этого объекта.|
||[свойства: { [клавиша: строка]](/javascript/api/excel/excel.entitycellvalue#excel-excel-entitycellvalue-properties-member)|Представляет свойства этого объекта и их метаданные.|
||[поставщик](/javascript/api/excel/excel.entitycellvalue#excel-excel-entitycellvalue-provider-member)|Представляет сведения, описывая службу, которая предоставила данные в этом .`EntityCellValue`|
||[referencedValues](/javascript/api/excel/excel.entitycellvalue#excel-excel-entitycellvalue-referencedvalues-member)|Представляет значения ячейки, которые ссылались в пределах `EntityCellValue.properties`.|
||[text](/javascript/api/excel/excel.entitycellvalue#excel-excel-entitycellvalue-text-member)|Представляет текст, показанный при отрисовывии ячейки с этим значением.|
||[тип](/javascript/api/excel/excel.entitycellvalue#excel-excel-entitycellvalue-type-member)|Представляет тип этого значения ячейки.|
|[EntityPropertyExtraProperties](/javascript/api/excel/excel.entitypropertyextraproperties)|[propertyMetadata](/javascript/api/excel/excel.entitypropertyextraproperties#excel-excel-entitypropertyextraproperties-propertymetadata-member)|Представляет метаданные об свойстве.|
|[EntityViewLayouts](/javascript/api/excel/excel.entityviewlayouts)|[карточка](/javascript/api/excel/excel.entityviewlayouts#excel-excel-entityviewlayouts-card-member)|Представляет макет этого объекта в представлении карты.|
|[FieldErrorCellValue](/javascript/api/excel/excel.fielderrorcellvalue)|[basicType](/javascript/api/excel/excel.fielderrorcellvalue#excel-excel-fielderrorcellvalue-basictype-member)|Представляет значение, которое будет возвращено ячейкой `Range.valueTypes` с этим значением.|
||[basicValue](/javascript/api/excel/excel.fielderrorcellvalue#excel-excel-fielderrorcellvalue-basicvalue-member)|Представляет значение, которое будет возвращено ячейкой `Range.values` с этим значением.|
||[errorSubType](/javascript/api/excel/excel.fielderrorcellvalue#excel-excel-fielderrorcellvalue-errorsubtype-member)|Представляет тип `FieldErrorCellValue`.|
||[errorType](/javascript/api/excel/excel.fielderrorcellvalue#excel-excel-fielderrorcellvalue-errortype-member)|Представляет тип `ErrorCellValue`.|
||[тип](/javascript/api/excel/excel.fielderrorcellvalue#excel-excel-fielderrorcellvalue-type-member)|Представляет тип этого значения ячейки.|
|[FormattedNumberCellValue](/javascript/api/excel/excel.formattednumbercellvalue)|[basicType](/javascript/api/excel/excel.formattednumbercellvalue#excel-excel-formattednumbercellvalue-basictype-member)|Представляет значение, которое будет возвращено ячейкой `Range.valueTypes` с этим значением.|
||[basicValue](/javascript/api/excel/excel.formattednumbercellvalue#excel-excel-formattednumbercellvalue-basicvalue-member)|Представляет значение, которое будет возвращено ячейкой `Range.values` с этим значением.|
||[numberFormat](/javascript/api/excel/excel.formattednumbercellvalue#excel-excel-formattednumbercellvalue-numberformat-member)|Возвращает строку формата номеров, которая используется для отображения этого значения.|
||[тип](/javascript/api/excel/excel.formattednumbercellvalue#excel-excel-formattednumbercellvalue-type-member)|Представляет тип этого значения ячейки.|
|[GettingDataErrorCellValue](/javascript/api/excel/excel.gettingdataerrorcellvalue)|[basicType](/javascript/api/excel/excel.gettingdataerrorcellvalue#excel-excel-gettingdataerrorcellvalue-basictype-member)|Представляет значение, которое будет возвращено ячейкой `Range.valueTypes` с этим значением.|
||[basicValue](/javascript/api/excel/excel.gettingdataerrorcellvalue#excel-excel-gettingdataerrorcellvalue-basicvalue-member)|Представляет значение, которое будет возвращено ячейкой `Range.values` с этим значением.|
||[errorType](/javascript/api/excel/excel.gettingdataerrorcellvalue#excel-excel-gettingdataerrorcellvalue-errortype-member)|Представляет тип `ErrorCellValue`.|
||[тип](/javascript/api/excel/excel.gettingdataerrorcellvalue#excel-excel-gettingdataerrorcellvalue-type-member)|Представляет тип этого значения ячейки.|
|[Identity](/javascript/api/excel/excel.identity)|[displayName](/javascript/api/excel/excel.identity#excel-excel-identity-displayname-member)|Представляет отображаемое имя пользователя.|
||[email](/javascript/api/excel/excel.identity#excel-excel-identity-email-member)|Представляет электронный адрес пользователя.|
||[id](/javascript/api/excel/excel.identity#excel-excel-identity-id-member)|Представляет уникальный ID пользователя.|
|[IdentityCollection](/javascript/api/excel/excel.identitycollection)|[add(assignee: Identity)](/javascript/api/excel/excel.identitycollection#excel-excel-identitycollection-add-member(1))|Добавляет идентификатор пользователя в коллекцию.|
||[clear()](/javascript/api/excel/excel.identitycollection#excel-excel-identitycollection-clear-member(1))|Удаляет все идентификаторы пользователей из коллекции.|
||[getCount()](/javascript/api/excel/excel.identitycollection#excel-excel-identitycollection-getcount-member(1))|Возвращает число элементов в коллекции.|
||[getItemAt(index: number)](/javascript/api/excel/excel.identitycollection#excel-excel-identitycollection-getitemat-member(1))|Получает удостоверение пользователя документа с помощью индекса в коллекции.|
||[items](/javascript/api/excel/excel.identitycollection#excel-excel-identitycollection-items-member)|Получает загруженные дочерние элементы в этой коллекции.|
||[remove(assignee: Identity)](/javascript/api/excel/excel.identitycollection#excel-excel-identitycollection-remove-member(1))|Удаляет удостоверение пользователя из коллекции.|
|[IdentityEntity](/javascript/api/excel/excel.identityentity)|[displayName](/javascript/api/excel/excel.identityentity#excel-excel-identityentity-displayname-member)|Представляет отображаемое имя пользователя.|
||[email](/javascript/api/excel/excel.identityentity#excel-excel-identityentity-email-member)|Представляет электронный адрес пользователя.|
||[id](/javascript/api/excel/excel.identityentity#excel-excel-identityentity-id-member)|Представляет уникальный ID пользователя.|
|[LinkedDataType](/javascript/api/excel/excel.linkeddatatype)|[dataProvider](/javascript/api/excel/excel.linkeddatatype#excel-excel-linkeddatatype-dataprovider-member)|Имя поставщика данных для связанного типа данных.|
||[lastRefreshed](/javascript/api/excel/excel.linkeddatatype#excel-excel-linkeddatatype-lastrefreshed-member)|Дата и время локального часового пояса с момента открытия книги при последнем обновлении связанного типа данных.|
||[name](/javascript/api/excel/excel.linkeddatatype#excel-excel-linkeddatatype-name-member)|Имя связанного типа данных.|
||[periodicRefreshInterval](/javascript/api/excel/excel.linkeddatatype#excel-excel-linkeddatatype-periodicrefreshinterval-member)|Частота в секундах, `refreshMode` при которой тип связанных данных обновляется, если установлено "Периодическое".|
||[refreshMode](/javascript/api/excel/excel.linkeddatatype#excel-excel-linkeddatatype-refreshmode-member)|Механизм получения данных для связанного типа данных.|
||[requestRefresh()](/javascript/api/excel/excel.linkeddatatype#excel-excel-linkeddatatype-requestrefresh-member(1))|Делает запрос на обновление связанного типа данных.|
||[requestSetRefreshMode(refreshMode: Excel. LinkedDataTypeRefreshMode)](/javascript/api/excel/excel.linkeddatatype#excel-excel-linkeddatatype-requestsetrefreshmode-member(1))|Делает запрос на изменение режима обновления для этого связанного типа данных.|
||[serviceId](/javascript/api/excel/excel.linkeddatatype#excel-excel-linkeddatatype-serviceid-member)|Уникальный ID связанного типа данных.|
||[supportedRefreshModes](/javascript/api/excel/excel.linkeddatatype#excel-excel-linkeddatatype-supportedrefreshmodes-member)|Возвращает массив со всеми режимами обновления, поддерживаемыми типом связанных данных.|
|[LinkedDataTypeAddedEventArgs](/javascript/api/excel/excel.linkeddatatypeaddedeventargs)|[serviceId](/javascript/api/excel/excel.linkeddatatypeaddedeventargs#excel-excel-linkeddatatypeaddedeventargs-serviceid-member)|Уникальный ID нового типа связанных данных.|
||[источник](/javascript/api/excel/excel.linkeddatatypeaddedeventargs#excel-excel-linkeddatatypeaddedeventargs-source-member)|Получает источник события.|
||[type](/javascript/api/excel/excel.linkeddatatypeaddedeventargs#excel-excel-linkeddatatypeaddedeventargs-type-member)|Получает тип события.|
|[LinkedDataTypeCollection](/javascript/api/excel/excel.linkeddatatypecollection)|[getCount()](/javascript/api/excel/excel.linkeddatatypecollection#excel-excel-linkeddatatypecollection-getcount-member(1))|Получает количество связанных типов данных в коллекции.|
||[getItem(key: number)](/javascript/api/excel/excel.linkeddatatypecollection#excel-excel-linkeddatatypecollection-getitem-member(1))|Получает связанный тип данных по ID службы.|
||[getItemAt(index: number)](/javascript/api/excel/excel.linkeddatatypecollection#excel-excel-linkeddatatypecollection-getitemat-member(1))|Получает связанный тип данных по индексу в коллекции.|
||[getItemOrNullObject(key: number)](/javascript/api/excel/excel.linkeddatatypecollection#excel-excel-linkeddatatypecollection-getitemornullobject-member(1))|Получает связанный тип данных по ID.|
||[items](/javascript/api/excel/excel.linkeddatatypecollection#excel-excel-linkeddatatypecollection-items-member)|Получает загруженные дочерние элементы в этой коллекции.|
||[requestRefreshAll()](/javascript/api/excel/excel.linkeddatatypecollection#excel-excel-linkeddatatypecollection-requestrefreshall-member(1))|Делает запрос на обновление всех связанных типов данных в коллекции.|
|[LinkedEntityCellValue](/javascript/api/excel/excel.linkedentitycellvalue)|[basicType](/javascript/api/excel/excel.linkedentitycellvalue#excel-excel-linkedentitycellvalue-basictype-member)|Представляет значение, которое будет возвращено ячейкой `Range.valueTypes` с этим значением.|
||[basicValue](/javascript/api/excel/excel.linkedentitycellvalue#excel-excel-linkedentitycellvalue-basicvalue-member)|Представляет значение, которое будет возвращено ячейкой `Range.values` с этим значением.|
||[cardLayout](/javascript/api/excel/excel.linkedentitycellvalue#excel-excel-linkedentitycellvalue-cardlayout-member)|Представляет макет этого связанного объекта в представлении карты.|
||[id](/javascript/api/excel/excel.linkedentitycellvalue#excel-excel-linkedentitycellvalue-id-member)|Представляет источник службы, который предоставил сведения в этом значении.|
||[свойства: { [клавиша: строка]: CellValue & {propertyMetadata](/javascript/api/excel/excel.linkedentitycellvalue#excel-excel-linkedentitycellvalue-properties-member)|Представляет свойства этой связанной сущности и их метаданных.|
||[propertyMetadata](/javascript/api/excel/excel.linkedentitycellvalue#excel-excel-linkedentitycellvalue-propertymetadata-member)||
||[поставщик](/javascript/api/excel/excel.linkedentitycellvalue#excel-excel-linkedentitycellvalue-provider-member)|Представляет сведения, описывая службу, которая предоставила данные в этом .`LinkedEntityCellValue`|
||[text](/javascript/api/excel/excel.linkedentitycellvalue#excel-excel-linkedentitycellvalue-text-member)|Представляет текст, показанный при отрисовывии ячейки с этим значением.|
||[тип](/javascript/api/excel/excel.linkedentitycellvalue#excel-excel-linkedentitycellvalue-type-member)|Представляет тип этого значения ячейки.|
|[LinkedEntityId](/javascript/api/excel/excel.linkedentityid)|[культура](/javascript/api/excel/excel.linkedentityid#excel-excel-linkedentityid-culture-member)|Представляет, какая языковая культура была использована для создания этого `CellValue`.|
||[domainId](/javascript/api/excel/excel.linkedentityid#excel-excel-linkedentityid-domainid-member)|Представляет домен, специфический для службы, используемой для создания `CellValue`.|
||[entityId](/javascript/api/excel/excel.linkedentityid#excel-excel-linkedentityid-entityid-member)|Представляет идентификатор, специфический для службы, используемой для создания `CellValue`.|
||[serviceId](/javascript/api/excel/excel.linkedentityid#excel-excel-linkedentityid-serviceid-member)|Представляет, какая служба была использована для создания `CellValue`.|
|[NameErrorCellValue](/javascript/api/excel/excel.nameerrorcellvalue)|[basicType](/javascript/api/excel/excel.nameerrorcellvalue#excel-excel-nameerrorcellvalue-basictype-member)|Представляет значение, которое будет возвращено ячейкой `Range.valueTypes` с этим значением.|
||[basicValue](/javascript/api/excel/excel.nameerrorcellvalue#excel-excel-nameerrorcellvalue-basicvalue-member)|Представляет значение, которое будет возвращено ячейкой `Range.values` с этим значением.|
||[errorType](/javascript/api/excel/excel.nameerrorcellvalue#excel-excel-nameerrorcellvalue-errortype-member)|Представляет тип `ErrorCellValue`.|
||[тип](/javascript/api/excel/excel.nameerrorcellvalue#excel-excel-nameerrorcellvalue-type-member)|Представляет тип этого значения ячейки.|
|[NamedItem](/javascript/api/excel/excel.nameditem)|[valueAsJson](/javascript/api/excel/excel.nameditem#excel-excel-nameditem-valueasjson-member)|Представление JSON значений в этом элементе с именем.|
||[valueAsJsonLocal](/javascript/api/excel/excel.nameditem#excel-excel-nameditem-valueasjsonlocal-member)|Представление JSON значений в этом элементе с именем.|
|[NamedItemArrayValues](/javascript/api/excel/excel.nameditemarrayvalues)|[valuesAsJson](/javascript/api/excel/excel.nameditemarrayvalues#excel-excel-nameditemarrayvalues-valuesasjson-member)|Представление JSON значений в этом массиве элементов с именем.|
||[valuesAsJsonLocal](/javascript/api/excel/excel.nameditemarrayvalues#excel-excel-nameditemarrayvalues-valuesasjsonlocal-member)|Представление JSON значений в этом массиве элементов с именем.|
|[NamedSheetViewCollection](/javascript/api/excel/excel.namedsheetviewcollection)|[getItemOrNullObject(key: string)](/javascript/api/excel/excel.namedsheetviewcollection#excel-excel-namedsheetviewcollection-getitemornullobject-member(1))|Получает представление листа с его именем.|
|[NotAvailableErrorCellValue](/javascript/api/excel/excel.notavailableerrorcellvalue)|[basicType](/javascript/api/excel/excel.notavailableerrorcellvalue#excel-excel-notavailableerrorcellvalue-basictype-member)|Представляет значение, которое будет возвращено ячейкой `Range.valueTypes` с этим значением.|
||[basicValue](/javascript/api/excel/excel.notavailableerrorcellvalue#excel-excel-notavailableerrorcellvalue-basicvalue-member)|Представляет значение, которое будет возвращено ячейкой `Range.values` с этим значением.|
||[errorType](/javascript/api/excel/excel.notavailableerrorcellvalue#excel-excel-notavailableerrorcellvalue-errortype-member)|Представляет тип `ErrorCellValue`.|
||[тип](/javascript/api/excel/excel.notavailableerrorcellvalue#excel-excel-notavailableerrorcellvalue-type-member)|Представляет тип этого значения ячейки.|
|[NullErrorCellValue](/javascript/api/excel/excel.nullerrorcellvalue)|[basicType](/javascript/api/excel/excel.nullerrorcellvalue#excel-excel-nullerrorcellvalue-basictype-member)|Представляет значение, которое будет возвращено ячейкой `Range.valueTypes` с этим значением.|
||[basicValue](/javascript/api/excel/excel.nullerrorcellvalue#excel-excel-nullerrorcellvalue-basicvalue-member)|Представляет значение, которое будет возвращено ячейкой `Range.values` с этим значением.|
||[errorType](/javascript/api/excel/excel.nullerrorcellvalue#excel-excel-nullerrorcellvalue-errortype-member)|Представляет тип `ErrorCellValue`.|
||[тип](/javascript/api/excel/excel.nullerrorcellvalue#excel-excel-nullerrorcellvalue-type-member)|Представляет тип этого значения ячейки.|
|[NumErrorCellValue](/javascript/api/excel/excel.numerrorcellvalue)|[basicType](/javascript/api/excel/excel.numerrorcellvalue#excel-excel-numerrorcellvalue-basictype-member)|Представляет значение, которое будет возвращено ячейкой `Range.valueTypes` с этим значением.|
||[basicValue](/javascript/api/excel/excel.numerrorcellvalue#excel-excel-numerrorcellvalue-basicvalue-member)|Представляет значение, которое будет возвращено ячейкой `Range.values` с этим значением.|
||[errorType](/javascript/api/excel/excel.numerrorcellvalue#excel-excel-numerrorcellvalue-errortype-member)|Представляет тип `ErrorCellValue`.|
||[тип](/javascript/api/excel/excel.numerrorcellvalue#excel-excel-numerrorcellvalue-type-member)|Представляет тип этого значения ячейки.|
|[PivotLayout](/javascript/api/excel/excel.pivotlayout)|[getCell(dataHierarchy: DataPivotHierarchy \| string, rowItems: Array<PivotItem \| string>, columnItems: Array<PivotItem \| string>)](/javascript/api/excel/excel.pivotlayout#excel-excel-pivotlayout-getcell-member(1))|Получает уникальную ячейку в сводной таблице на основе иерархии данных и элементов строк и столбцов соответствующих иерархий.|
||[pivotStyle](/javascript/api/excel/excel.pivotlayout#excel-excel-pivotlayout-pivotstyle-member)|Стиль, примененный к PivotTable.|
||[setStyle(style: string \| PivotTableStyle \| BuiltInPivotTableStyle)](/javascript/api/excel/excel.pivotlayout#excel-excel-pivotlayout-setstyle-member(1))|Задает стиль, применяемый к PivotTable.|
|[PivotTable](/javascript/api/excel/excel.pivottable)|[getDataSourceString()](/javascript/api/excel/excel.pivottable#excel-excel-pivottable-getdatasourcestring-member(1))|Возвращает представление строк источника данных для PivotTable.|
||[getDataSourceType()](/javascript/api/excel/excel.pivottable#excel-excel-pivottable-getdatasourcetype-member(1))|Получает тип источника данных для PivotTable.|
|[PivotTableScopedCollection](/javascript/api/excel/excel.pivottablescopedcollection)|[getFirstOrNullObject()](/javascript/api/excel/excel.pivottablescopedcollection#excel-excel-pivottablescopedcollection-getfirstornullobject-member(1))|Получает первый pivotTable в коллекции.|
|[PlaceholderErrorCellValue](/javascript/api/excel/excel.placeholdererrorcellvalue)|[basicType](/javascript/api/excel/excel.placeholdererrorcellvalue#excel-excel-placeholdererrorcellvalue-basictype-member)|Представляет значение, которое будет возвращено ячейкой `Range.valueTypes` с этим значением.|
||[basicValue](/javascript/api/excel/excel.placeholdererrorcellvalue#excel-excel-placeholdererrorcellvalue-basicvalue-member)|Представляет значение, которое будет возвращено ячейкой `Range.values` с этим значением.|
||[errorType](/javascript/api/excel/excel.placeholdererrorcellvalue#excel-excel-placeholdererrorcellvalue-errortype-member)|Представляет тип `ErrorCellValue`.|
||[конечный объект](/javascript/api/excel/excel.placeholdererrorcellvalue#excel-excel-placeholdererrorcellvalue-target-member)|`PlaceholderErrorCellValue` используется во время обработки при загрузке данных.|
||[тип](/javascript/api/excel/excel.placeholdererrorcellvalue#excel-excel-placeholdererrorcellvalue-type-member)|Представляет тип этого значения ячейки.|
|[Range](/javascript/api/excel/excel.range)|[getDependents()](/javascript/api/excel/excel.range#excel-excel-range-getdependents-member(1))|Возвращает объект, `WorkbookRangeAreas` представляющего диапазон, содержащий все иждивенцы ячейки в одной и той же таблице или в нескольких таблицах.|
||[valuesAsJson](/javascript/api/excel/excel.range#excel-excel-range-valuesasjson-member)|Представление JSON значений в ячейках этого диапазона.|
||[valuesAsJsonLocal](/javascript/api/excel/excel.range#excel-excel-range-valuesasjsonlocal-member)|Представление JSON значений в ячейках этого диапазона.|
|[RangeView](/javascript/api/excel/excel.rangeview)|[valuesAsJson](/javascript/api/excel/excel.rangeview#excel-excel-rangeview-valuesasjson-member)|Представление JSON значений в ячейках этого диапазона.|
||[valuesAsJsonLocal](/javascript/api/excel/excel.rangeview#excel-excel-rangeview-valuesasjsonlocal-member)|Представление JSON значений в ячейках этого диапазона.|
|[RefErrorCellValue](/javascript/api/excel/excel.referrorcellvalue)|[basicType](/javascript/api/excel/excel.referrorcellvalue#excel-excel-referrorcellvalue-basictype-member)|Представляет значение, которое будет возвращено ячейкой `Range.valueTypes` с этим значением.|
||[basicValue](/javascript/api/excel/excel.referrorcellvalue#excel-excel-referrorcellvalue-basicvalue-member)|Представляет значение, которое будет возвращено ячейкой `Range.values` с этим значением.|
||[errorSubType](/javascript/api/excel/excel.referrorcellvalue#excel-excel-referrorcellvalue-errorsubtype-member)|Представляет тип `RefErrorCellValue`.|
||[errorType](/javascript/api/excel/excel.referrorcellvalue#excel-excel-referrorcellvalue-errortype-member)|Представляет тип `ErrorCellValue`.|
||[тип](/javascript/api/excel/excel.referrorcellvalue#excel-excel-referrorcellvalue-type-member)|Представляет тип этого значения ячейки.|
|[ReferenceCellValue](/javascript/api/excel/excel.referencecellvalue)|[basicType](/javascript/api/excel/excel.referencecellvalue#excel-excel-referencecellvalue-basictype-member)|Представляет значение, которое будет возвращено ячейкой `Range.valueTypes` с этим значением.|
||[basicValue](/javascript/api/excel/excel.referencecellvalue#excel-excel-referencecellvalue-basicvalue-member)|Представляет значение, которое будет возвращено ячейкой `Range.values` с этим значением.|
||[reference](/javascript/api/excel/excel.referencecellvalue#excel-excel-referencecellvalue-reference-member)|Представляет индекс в свойствах `referencedValues` значений ячейки, таких как `EntityCellValue` и `ArrayCellValue`.|
||[тип](/javascript/api/excel/excel.referencecellvalue#excel-excel-referencecellvalue-type-member)|Представляет тип этого значения ячейки.|
|[RefreshModeChangedEventArgs](/javascript/api/excel/excel.refreshmodechangedeventargs)|[refreshMode](/javascript/api/excel/excel.refreshmodechangedeventargs#excel-excel-refreshmodechangedeventargs-refreshmode-member)|Режим обновления связанного типа данных.|
||[serviceId](/javascript/api/excel/excel.refreshmodechangedeventargs#excel-excel-refreshmodechangedeventargs-serviceid-member)|Уникальный ID объекта, режим обновления которого был изменен.|
||[источник](/javascript/api/excel/excel.refreshmodechangedeventargs#excel-excel-refreshmodechangedeventargs-source-member)|Получает источник события.|
||[type](/javascript/api/excel/excel.refreshmodechangedeventargs#excel-excel-refreshmodechangedeventargs-type-member)|Получает тип события.|
|[RefreshRequestCompletedEventArgs](/javascript/api/excel/excel.refreshrequestcompletedeventargs)|[обновлено](/javascript/api/excel/excel.refreshrequestcompletedeventargs#excel-excel-refreshrequestcompletedeventargs-refreshed-member)|Указывает, был ли запрос на обновление успешным.|
||[serviceId](/javascript/api/excel/excel.refreshrequestcompletedeventargs#excel-excel-refreshrequestcompletedeventargs-serviceid-member)|Уникальный ID объекта, запрос на обновление которого был завершен.|
||[источник](/javascript/api/excel/excel.refreshrequestcompletedeventargs#excel-excel-refreshrequestcompletedeventargs-source-member)|Получает источник события.|
||[type](/javascript/api/excel/excel.refreshrequestcompletedeventargs#excel-excel-refreshrequestcompletedeventargs-type-member)|Получает тип события.|
||[предупреждения](/javascript/api/excel/excel.refreshrequestcompletedeventargs#excel-excel-refreshrequestcompletedeventargs-warnings-member)|Массив, содержащий все предупреждения, созданные из запроса на обновление.|
|[RootReferenceCellValue](/javascript/api/excel/excel.rootreferencecellvalue)|[basicType](/javascript/api/excel/excel.rootreferencecellvalue#excel-excel-rootreferencecellvalue-basictype-member)|Представляет значение, которое будет возвращено ячейкой `Range.valueTypes` с этим значением.|
||[basicValue](/javascript/api/excel/excel.rootreferencecellvalue#excel-excel-rootreferencecellvalue-basicvalue-member)|Представляет значение, которое будет возвращено ячейкой `Range.values` с этим значением.|
||[тип](/javascript/api/excel/excel.rootreferencecellvalue#excel-excel-rootreferencecellvalue-type-member)|Представляет тип этого значения ячейки.|
|[Shape](/javascript/api/excel/excel.shape)|[displayName](/javascript/api/excel/excel.shape#excel-excel-shape-displayname-member)|Получает имя отображения фигуры.|
|[ShapeCollection](/javascript/api/excel/excel.shapecollection)|[addSvg(xml: string)](/javascript/api/excel/excel.shapecollection#excel-excel-shapecollection-addsvg-member(1))|Создает изображение SVG (масштабируемая векторная графика) из строки XML и добавляет его на лист.|
|[Slicer](/javascript/api/excel/excel.slicer)|[nameInFormula](/javascript/api/excel/excel.slicer#excel-excel-slicer-nameinformula-member)|Представляет имя среза, используемое в формуле.|
||[setStyle(style: string \| SlicerStyle \| BuiltInSlicerStyle)](/javascript/api/excel/excel.slicer#excel-excel-slicer-setstyle-member(1))|Задает стиль, примененный к срезу.|
||[slicerStyle](/javascript/api/excel/excel.slicer#excel-excel-slicer-slicerstyle-member)|Стиль, применяемый к срезу.|
|[SpillErrorCellValue](/javascript/api/excel/excel.spillerrorcellvalue)|[basicType](/javascript/api/excel/excel.spillerrorcellvalue#excel-excel-spillerrorcellvalue-basictype-member)|Представляет значение, которое будет возвращено ячейкой `Range.valueTypes` с этим значением.|
||[basicValue](/javascript/api/excel/excel.spillerrorcellvalue#excel-excel-spillerrorcellvalue-basicvalue-member)|Представляет значение, которое будет возвращено ячейкой `Range.values` с этим значением.|
||[errorSubType](/javascript/api/excel/excel.spillerrorcellvalue#excel-excel-spillerrorcellvalue-errorsubtype-member)|Представляет тип `SpillErrorCellValue`.|
||[errorType](/javascript/api/excel/excel.spillerrorcellvalue#excel-excel-spillerrorcellvalue-errortype-member)|Представляет тип `ErrorCellValue`.|
||[spilledColumns](/javascript/api/excel/excel.spillerrorcellvalue#excel-excel-spillerrorcellvalue-spilledcolumns-member)|Представляет количество столбцов, которые будут разливаться, если бы не было #SPILL! ошибка.|
||[spilledRows](/javascript/api/excel/excel.spillerrorcellvalue#excel-excel-spillerrorcellvalue-spilledrows-member)|Представляет количество строк, которые разлились бы, если бы не было #SPILL! ошибка.|
||[тип](/javascript/api/excel/excel.spillerrorcellvalue#excel-excel-spillerrorcellvalue-type-member)|Представляет тип этого значения ячейки.|
|[StringCellValue](/javascript/api/excel/excel.stringcellvalue)|[basicType](/javascript/api/excel/excel.stringcellvalue#excel-excel-stringcellvalue-basictype-member)|Представляет значение, которое будет возвращено ячейкой `Range.valueTypes` с этим значением.|
||[basicValue](/javascript/api/excel/excel.stringcellvalue#excel-excel-stringcellvalue-basicvalue-member)|Представляет значение, которое будет возвращено ячейкой `Range.values` с этим значением.|
||[тип](/javascript/api/excel/excel.stringcellvalue#excel-excel-stringcellvalue-type-member)|Представляет тип этого значения ячейки.|
|[Table](/javascript/api/excel/excel.table)|[clearStyle()](/javascript/api/excel/excel.table#excel-excel-table-clearstyle-member(1))|Изменяет таблицу для использования стиля таблицы по умолчанию.|
||[onFiltered](/javascript/api/excel/excel.table#excel-excel-table-onfiltered-member)|Возникает, когда фильтр применяется на определенной таблице.|
||[setStyle(style: string \| TableStyle \| BuiltInTableStyle)](/javascript/api/excel/excel.table#excel-excel-table-setstyle-member(1))|Задает стиль, примененный к таблице.|
||[tableStyle](/javascript/api/excel/excel.table#excel-excel-table-tablestyle-member)|Стиль, примененный к таблице.|
|[TableCollection](/javascript/api/excel/excel.tablecollection)|[onFiltered](/javascript/api/excel/excel.tablecollection#excel-excel-tablecollection-onfiltered-member)|Возникает, когда фильтр применяется на любой таблице в книге или в таблице.|
|[TableColumn](/javascript/api/excel/excel.tablecolumn)|[valuesAsJson](/javascript/api/excel/excel.tablecolumn#excel-excel-tablecolumn-valuesasjson-member)|Представление JSON значений в ячейках в этом столбце таблицы.|
||[valuesAsJsonLocal](/javascript/api/excel/excel.tablecolumn#excel-excel-tablecolumn-valuesasjsonlocal-member)|Представление JSON значений в ячейках в этом столбце таблицы.|
|[TableFilteredEventArgs](/javascript/api/excel/excel.tablefilteredeventargs)|[tableId](/javascript/api/excel/excel.tablefilteredeventargs#excel-excel-tablefilteredeventargs-tableid-member)|Получает ID таблицы, в которой применяется фильтр.|
||[type](/javascript/api/excel/excel.tablefilteredeventargs#excel-excel-tablefilteredeventargs-type-member)|Получает тип события.|
||[worksheetId](/javascript/api/excel/excel.tablefilteredeventargs#excel-excel-tablefilteredeventargs-worksheetid-member)|Получает ID таблицы, которая содержит таблицу.|
|[TableRow](/javascript/api/excel/excel.tablerow)|[valuesAsJson](/javascript/api/excel/excel.tablerow#excel-excel-tablerow-valuesasjson-member)|Представление JSON значений в ячейках в этой строке таблицы.|
||[valuesAsJsonLocal](/javascript/api/excel/excel.tablerow#excel-excel-tablerow-valuesasjsonlocal-member)|Представление JSON значений в ячейках в этой строке таблицы.|
|[ValueErrorCellValue](/javascript/api/excel/excel.valueerrorcellvalue)|[basicType](/javascript/api/excel/excel.valueerrorcellvalue#excel-excel-valueerrorcellvalue-basictype-member)|Представляет значение, которое будет возвращено ячейкой `Range.valueTypes` с этим значением.|
||[basicValue](/javascript/api/excel/excel.valueerrorcellvalue#excel-excel-valueerrorcellvalue-basicvalue-member)|Представляет значение, которое будет возвращено ячейкой `Range.values` с этим значением.|
||[errorSubType](/javascript/api/excel/excel.valueerrorcellvalue#excel-excel-valueerrorcellvalue-errorsubtype-member)|Представляет тип `ValueErrorCellValue`.|
||[errorType](/javascript/api/excel/excel.valueerrorcellvalue#excel-excel-valueerrorcellvalue-errortype-member)|Представляет тип `ErrorCellValue`.|
||[тип](/javascript/api/excel/excel.valueerrorcellvalue#excel-excel-valueerrorcellvalue-type-member)|Представляет тип этого значения ячейки.|
|[ValueTypeNotAvailableCellValue](/javascript/api/excel/excel.valuetypenotavailablecellvalue)|[basicType](/javascript/api/excel/excel.valuetypenotavailablecellvalue#excel-excel-valuetypenotavailablecellvalue-basictype-member)|Представляет значение, которое будет возвращено ячейкой `Range.valueTypes` с этим значением.|
||[basicValue](/javascript/api/excel/excel.valuetypenotavailablecellvalue#excel-excel-valuetypenotavailablecellvalue-basicvalue-member)|Представляет значение, которое будет возвращено ячейкой `Range.values` с этим значением.|
||[тип](/javascript/api/excel/excel.valuetypenotavailablecellvalue#excel-excel-valuetypenotavailablecellvalue-type-member)|Представляет тип этого значения ячейки.|
|[WebImageCellValue](/javascript/api/excel/excel.webimagecellvalue)|[address](/javascript/api/excel/excel.webimagecellvalue#excel-excel-webimagecellvalue-address-member)|Представляет URL-адрес, с которого будет загружено изображение.|
||[altText](/javascript/api/excel/excel.webimagecellvalue#excel-excel-webimagecellvalue-alttext-member)|Представляет альтернативный текст, который можно использовать в сценариях доступности для описания изображения.|
||[атрибуция](/javascript/api/excel/excel.webimagecellvalue#excel-excel-webimagecellvalue-attribution-member)|Представляет сведения о присвоении для описания исходных и лицензионных требований для использования этого изображения.|
||[basicType](/javascript/api/excel/excel.webimagecellvalue#excel-excel-webimagecellvalue-basictype-member)|Представляет значение, которое будет возвращено ячейкой `Range.valueTypes` с этим значением.|
||[basicValue](/javascript/api/excel/excel.webimagecellvalue#excel-excel-webimagecellvalue-basicvalue-member)|Представляет значение, которое будет возвращено ячейкой `Range.values` с этим значением.|
||[поставщик](/javascript/api/excel/excel.webimagecellvalue#excel-excel-webimagecellvalue-provider-member)|Представляет сведения, описывавшие сущность или физическое лицо, которое предоставило изображение.|
||[relatedImagesAddress](/javascript/api/excel/excel.webimagecellvalue#excel-excel-webimagecellvalue-relatedimagesaddress-member)|Представляет URL-адрес веб-страницы с изображениями, которые считаются связанными с этим `WebImageCellValue`.|
||[тип](/javascript/api/excel/excel.webimagecellvalue#excel-excel-webimagecellvalue-type-member)|Представляет тип этого значения ячейки.|
|[Workbook](/javascript/api/excel/excel.workbook)|[getLinkedEntityCellValue (linkedEntityCellValueId: LinkedEntityId)](/javascript/api/excel/excel.workbook#excel-excel-workbook-getlinkedentitycellvalue-member(1))|Возвращает на основе `LinkedEntityCellValue` предоставленного `LinkedEntityId`.|
||[linkedDataTypes](/javascript/api/excel/excel.workbook#excel-excel-workbook-linkeddatatypes-member)|Возвращает коллекцию связанных типов данных, которые являются частью книги.|
||[showPivotFieldList](/javascript/api/excel/excel.workbook#excel-excel-workbook-showpivotfieldlist-member)|Указывает, отображается ли область списка полей PivotTable на уровне книги.|
||[задачи](/javascript/api/excel/excel.workbook#excel-excel-workbook-tasks-member)|Возвращает коллекцию задач, присутствующих в книге.|
||[use1904DateSystem](/javascript/api/excel/excel.workbook#excel-excel-workbook-use1904datesystem-member)|Значение true, если в книге используется система дат 1904.|
|[Worksheet](/javascript/api/excel/excel.worksheet)|[onFiltered](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-onfiltered-member)|Возникает, когда фильтр применяется на определенном таблице.|
||[задачи](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-tasks-member)|Возвращает коллекцию задач, присутствующих в таблице.|
|[WorksheetCollection](/javascript/api/excel/excel.worksheetcollection)|[addFromBase64(base64File: string, sheetNamesToInsert?: string[], positionType?: Excel.WorksheetPositionType, relativeTo?: Worksheet \| string)](/javascript/api/excel/excel.worksheetcollection#excel-excel-worksheetcollection-addfrombase64-member(1))|Вставляет указанные листы книги в текущую книгу.|
||[onFiltered](/javascript/api/excel/excel.worksheetcollection#excel-excel-worksheetcollection-onfiltered-member)|Возникает при применении любого фильтра листа в книге.|
|[WorksheetFilteredEventArgs](/javascript/api/excel/excel.worksheetfilteredeventargs)|[type](/javascript/api/excel/excel.worksheetfilteredeventargs#excel-excel-worksheetfilteredeventargs-type-member)|Получает тип события.|
||[worksheetId](/javascript/api/excel/excel.worksheetfilteredeventargs#excel-excel-worksheetfilteredeventargs-worksheetid-member)|Получает ID таблицы, в которой применяется фильтр.|
|[WorksheetProtection](/javascript/api/excel/excel.worksheetprotection)|[allowEditRanges](/javascript/api/excel/excel.worksheetprotection#excel-excel-worksheetprotection-alloweditranges-member)|Указывает объект `AllowEditRangeCollection` , найденный в этом документе.|
||[canPauseProtection](/javascript/api/excel/excel.worksheetprotection#excel-excel-worksheetprotection-canpauseprotection-member)|Указывает, можно ли приостановить защиту для этого таблицы.|
||[checkPassword(password?: string)](/javascript/api/excel/excel.worksheetprotection#excel-excel-worksheetprotection-checkpassword-member(1))|Указывает, можно ли использовать пароль для разблокировки защиты таблицы.|
||[isPasswordProtected](/javascript/api/excel/excel.worksheetprotection#excel-excel-worksheetprotection-ispasswordprotected-member)|Указывает, защищен ли лист паролем.|
||[isPaused](/javascript/api/excel/excel.worksheetprotection#excel-excel-worksheetprotection-ispaused-member)|Указывает, приостановлена ли защита таблицы.|
||[pauseProtection(password?: string)](/javascript/api/excel/excel.worksheetprotection#excel-excel-worksheetprotection-pauseprotection-member(1))|Приостановка защиты таблиц для данного объекта таблицы для пользователя в заданном сеансе.|
||[resumeProtection()](/javascript/api/excel/excel.worksheetprotection#excel-excel-worksheetprotection-resumeprotection-member(1))|Возобновляет защиту таблиц для данного объекта таблицы для пользователя в заданном сеансе.|
||[savedOptions](/javascript/api/excel/excel.worksheetprotection#excel-excel-worksheetprotection-savedoptions-member)|Указывает параметры защиты, сохраненные в таблице.|
||[setPassword(password?: string)](/javascript/api/excel/excel.worksheetprotection#excel-excel-worksheetprotection-setpassword-member(1))|Изменяет пароль, связанный с объектом `WorksheetProtection` .|
||[updateOptions (параметры: Excel. WorksheetProtectionOptions)](/javascript/api/excel/excel.worksheetprotection#excel-excel-worksheetprotection-updateoptions-member(1))|Измените параметры защиты таблиц, связанные с объектом `WorksheetProtection` .|
|[WorksheetProtectionChangedEventArgs](/javascript/api/excel/excel.worksheetprotectionchangedeventargs)|[allowEditRangesChanged](/javascript/api/excel/excel.worksheetprotectionchangedeventargs#excel-excel-worksheetprotectionchangedeventargs-alloweditrangeschanged-member)|Указывает, изменились ли `AllowEditRange` какие-либо объекты.|
||[protectionOptionsChanged](/javascript/api/excel/excel.worksheetprotectionchangedeventargs#excel-excel-worksheetprotectionchangedeventargs-protectionoptionschanged-member)|Указывает, изменились ли `WorksheetProtectionOptions` изменения.|
||[sheetPasswordChanged](/javascript/api/excel/excel.worksheetprotectionchangedeventargs#excel-excel-worksheetprotectionchangedeventargs-sheetpasswordchanged-member)|Указывает, изменился ли пароль таблицы.|
