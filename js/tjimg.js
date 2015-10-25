  function switnewspbq(tjimg_a,tjimg_b,tjimg) {
    for(i=1; i <7; i++) {
      if (i==tjimg) {
        document.getElementById(tjimg_a+i).className="tjimg_menuOn";document.getElementById(tjimg_b+i).className="incthj_conten";}
      else {
		  if (i==6) {
             document.getElementById(tjimg_a+i).className="tjimg_menuNo";document.getElementById(tjimg_b+i).className="incthj_conten_none";}
		  else {
             document.getElementById(tjimg_a+i).className="tjimg_menuNo";document.getElementById(tjimg_b+i).className="incthj_conten_none";}
    }
  }
}