function reminderDailyTodo(){
  let dailyReminder = new NotionReminder('Daily Task');
  dailyReminder.sendReminder();
}

function reminderITTodo(){
  let dailyReminder = new NotionReminder('IT Study');
  dailyReminder.sendReminder();
}

function reminderCardTodo(){
  let dailyReminder = new NotionReminder('Card Management');
  dailyReminder.sendReminder();
}

function reminderCarTodo(){
  let dailyReminder = new NotionReminder('Car Management');
  dailyReminder.sendReminder();
}

function reminderAssetTodo(){
  let dailyReminder = new NotionReminder('Asset Management');
  dailyReminder.sendReminder();
}