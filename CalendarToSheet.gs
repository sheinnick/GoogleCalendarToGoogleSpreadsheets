// http://sheinnikita.ru/googleappscrypt/export-and-analize-events-from-google-calendar-in-spreadsheets.html
// https://github.com/sheinnick/GoogleCalendarToGoogleSpreadsheets
// https://www.facebook.com/shein.nikita

function CalToSheet() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('source');         //берем лист source
  var cal = CalendarApp.getCalendarById(sheet.getRange(9, 1).getValue());              //берем аккаунт из A9
  var startDate = new Date(sheet.getRange(3, 1).getValue());                           //берем дату начала периода из A3
  var endDate = new Date(sheet.getRange(6, 1).getValue());                             //берем дату начала периода из A6  
  var dinner = sheet.getRange(11, 1).getValue();                                       //берём название которое надо ставить со знаком минус, обед time 
  var dinnerColor = sheet.getRange(14, 1).getValue();                                  //берём цвет в который покрасим обед. цвета тут https://developers.google.com/apps-script/reference/calendar/event-color
       // var dinnerColorType = typeof(dinnerColor)

  var events = cal.getEvents(startDate, endDate);
  var num = events.length;
  if (num > 0) {
    var eventsArray = []
    for (var i = 0; i < num; i++)     //собираем массив с инфой о событиях
    {
      var event = []
      //если это обед то делаем его значение минус, чтобы он вычитался из рабочего времени и меняем его цвет на MAUVE	Enum	 Mauve ("3").
      if (events[i].getTitle() == dinner) {
        event.push(events[i].getTitle());            //название
        event.push(events[i].getStartTime());        //начало
        event.push(events[i].getEndTime());          //конец
        event.push(-(events[i].getEndTime() - events[i].getStartTime()) / 3600 / 1000);   //считаем длительность в часах
        event.push(events[i].getAllTagKeys());       //теги
//        event.push(events[i].setColor(eval("CalendarApp.EventColor."+String(dinnerColor)))); //ENUM цвета
          event.push(events[i].setColor(dinnerColor.toString())); //номер цвета как строка
        event.push(events[i].getMyStatus());         //статус
      } else {
        event.push(events[i].getTitle());            //название
        event.push(events[i].getStartTime());        //начало
        event.push(events[i].getEndTime());          //конец
        event.push((events[i].getEndTime() - events[i].getStartTime()) / 3600 / 1000);   //считаем длительность в часах
        event.push(events[i].getAllTagKeys());       //теги
        event.push(events[i].getColor());            //номер цвета
        event.push(events[i].getMyStatus());         //статус
      }

      if (events[i].isRecurringEvent() == true)    //повторяющееся или нет
      {
        event.push('recuring');
      } else { event.push('') };

      if (events[i].isAllDayEvent() == true)       //на весь день или нет
      {
        event.push('AllDay');
      } else { event.push('') };

      event.push(events[i].getVisibility());       //видимость

      eventsArray.push(event)
    };

    sheet.getRange("source!C2:M").clearContent(); //очищаем С2:M от предыдущих импортов

    sheet.getRange(2, 3, eventsArray.length, eventsArray[0].length).setValues(eventsArray)  //записываем результат на страницу
  }
}
