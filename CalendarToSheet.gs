// sheinnikita.ru
// https://www.facebook.com/shein.nikita
// https://github.com/sheinnick/GoogleCalendarToGoogleSpreadsheets

function CalToSheet()
{ 
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('source');         //берем лист source
  var cal = CalendarApp.getCalendarById(sheet.getRange(9,1).getValue());              //берем аккаунт из A9
  var startDate = new Date(sheet.getRange(3,1).getValue());                           //берем дату начала периода из A3
  var endDate = new Date(sheet.getRange(6,1).getValue());                             //берем дату начала периода из A6  
  
  var events = cal.getEvents(startDate, endDate);
  var num= events.length;
  if (num > 0)
  {
    var eventsArray = []
    for (var i=0; i<num; i++)     //собираем массив с инфой о событиях
      {
        var event = []
        event.push(events[i].getTitle());            //название
        event.push(events[i].getStartTime());        //начало
        event.push(events[i].getEndTime());          //конец
        event.push((events[i].getEndTime()-events[i].getStartTime())/3600/1000);   //считаем длительность в часах
        event.push(events[i].getAllTagKeys());       //теги
        event.push(events[i].getColor());            //номер цвета
        event.push(events[i].getMyStatus());         //статус
        
        if (events[i].isRecurringEvent() == true)    //повторяющееся или нет
         {
           event.push('recuring');
         }else {event.push('')};
        
        if (events[i].isAllDayEvent() == true)       //на весь день или нет
         {
           event.push('AllDay');
         } else {event.push('')};
      
        event.push(events[i].getVisibility());       //видимость
        
        eventsArray.push(event)
    };

    sheet.getRange("source!C2:M").clearContent(); //очищаем С2:M от предыдущих импортов
     
    sheet.getRange(2, 3, eventsArray.length, eventsArray[0].length).setValues(eventsArray)  //записываем результат на страницу
  }
}
