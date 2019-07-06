function CalToSheet()
{ 
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('source');         //берем лист source
  var cal = CalendarApp.getCalendarById(sheet.getRange(9,1).getValue());              //берем аккаунт из A9
  var startDate = new Date(sheet.getRange(3,1).getValue());                           //берем дату начала периода из A3
  var endDate = new Date(sheet.getRange(6,1).getValue());                             //берем дату начала периода из A6  
  
  //очищаем С2:M от предыдущих импортов
  sheet.getRange("source!C2:M").clearContent();
  
  var events = cal.getEvents(startDate, endDate);
  var num= events.length;
  if (num > 0)
  {
    for (var i=0; i<num; i++)
      {
        sheet.getRange(i+2, 3).setValue(events[i].getTitle());            //название
        sheet.getRange(i+2, 4).setValue(events[i].getStartTime());        //начало
        sheet.getRange(i+2, 5).setValue(events[i].getEndTime());          //конец
        sheet.getRange(i+2, 6).setValue((events[i].getEndTime()-events[i].getStartTime())/3600/1000);   //считаем длительность в часах
        sheet.getRange(i+2, 7).setValue(events[i].getAllTagKeys());       //теги
        sheet.getRange(i+2, 8).setValue(events[i].getColor());            //номер цвета
        sheet.getRange(i+2, 9).setValue(events[i].getMyStatus());         //статус
        
        if (events[i].isRecurringEvent() == true)                         //повторяющееся или нет
         {
           sheet.getRange(i+2, 10).setValue('recuring');
         };
        
        if (events[i].isAllDayEvent() == true)                            //на весь день или нет
         {
           sheet.getRange(i+2, 11).setValue('AllDay');
         };
      
        sheet.getRange(i+2, 12).setValue(events[i].getVisibility());       //видимость
    }
  }
}
