&НаСервере
Функция ПолучитьДанныеAD(ЛогинAD) Экспорт

    Результат = Новый Структура;

    Результат.Вставить("Найден", Ложь);
    Результат.Вставить("Имя", "");
    Результат.Вставить("Фамилия", "");
    Результат.Вставить("Отчество", "");
    Результат.Вставить("Должность", "");
    Результат.Вставить("Подразделение", "");
    Результат.Вставить("Организация", "");
    Результат.Вставить("Заблокирован", Ложь);

    ЛогинAD = СокрЛП(ЛогинAD);

    Если ПустаяСтрока(ЛогинAD) Тогда
        Возврат Результат;
    КонецЕсли;

    Попытка

        Connection = Новый COMОбъект("ADODB.Connection");
        Connection.Provider = "ADsDSOObject";
        Connection.Open("Active Directory Provider");

        Command = Новый COMОбъект("ADODB.Command");
        Command.ActiveConnection = Connection;

        RootDSE = Новый COMОбъект("ADSI.LDAP");
        RootDSE = RootDSE.OpenLDAPObject("LDAP://RootDSE");

        DefaultNamingContext = RootDSE.Get("defaultNamingContext");

        ТекстЗапроса =
        "<LDAP://" + DefaultNamingContext + ">;" +
        "(&(objectCategory=person)(objectClass=user)(sAMAccountName=" + ЛогинAD + "));" +
        "givenName,sn,middleName,title,department,company,userAccountControl;" +
        "subtree";

        Command.CommandText = ТекстЗапроса;

        RecordSet = Command.Execute();

        Если НЕ RecordSet.EOF Тогда

            Результат.Найден = Истина;

            Если ЗначениеЗаполнено(RecordSet.Fields("givenName").Value) Тогда
                Результат.Имя = Строка(RecordSet.Fields("givenName").Value);
            КонецЕсли;

            Если ЗначениеЗаполнено(RecordSet.Fields("sn").Value) Тогда
                Результат.Фамилия = Строка(RecordSet.Fields("sn").Value);
            КонецЕсли;

            Если ЗначениеЗаполнено(RecordSet.Fields("middleName").Value) Тогда
                Результат.Отчество = Строка(RecordSet.Fields("middleName").Value);
            КонецЕсли;

            Если ЗначениеЗаполнено(RecordSet.Fields("title").Value) Тогда
                Результат.Должность = Строка(RecordSet.Fields("title").Value);
            КонецЕсли;

            Если ЗначениеЗаполнено(RecordSet.Fields("department").Value) Тогда
                Результат.Подразделение = Строка(RecordSet.Fields("department").Value);
            КонецЕсли;

            Если ЗначениеЗаполнено(RecordSet.Fields("company").Value) Тогда
                Результат.Организация = Строка(RecordSet.Fields("company").Value);
            КонецЕсли;

            UserAccountControl = 0;

            Если ЗначениеЗаполнено(RecordSet.Fields("userAccountControl").Value) Тогда
                UserAccountControl = Число(RecordSet.Fields("userAccountControl").Value);
            КонецЕсли;

            Результат.Заблокирован = ((UserAccountControl И 2) = 2);

        КонецЕсли;

        RecordSet.Close();
        Connection.Close();

    Исключение

        Сообщить("Ошибка Active Directory: " + ОписаниеОшибки());

    КонецПопытки;

    Возврат Результат;

КонецФункции