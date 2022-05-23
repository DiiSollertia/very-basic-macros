function onEdit(e) {
  try {
    var rng = e.range;
    Logger.log(e.triggerUid);
    if(rng.isBlank()){
      Logger.log('Cell deleted');
      rng.setBackground(null);
      rng.setNote(null);
      return null;
    }

    var clr = e.value;
    var firstComma = clr.indexOf(",");
    var secondComma = clr.indexOf(",",firstComma + 1)
    
    if(clr.length <= 11 && clr.length >= 5 && clr.count(",") == 2){
      rng.setNote(null);
      rng.setBackgroundRGB(clr.slice(0,firstComma),clr.slice(firstComma + 1, secondComma), clr.slice(secondComma + 1));
    }
    else{
      rng.setNote('Invalid Format: Not a string of valid length.');
    }
  }
  catch(err){
    Logger.log('Failed with error %s', err.message);
  }
}