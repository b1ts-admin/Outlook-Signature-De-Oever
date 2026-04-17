// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

function get_template_A_str(user_info)
{
  let str = "";
  if (is_valid_data(user_info.greeting))
  {
    str += "<p>" + user_info.greeting + "</p>";
  }

  str += "<p style=\"color:#004259;font-size:medium;\">";
  str += "<strong>" + user_info.name + "</strong>";
  str += "<br/>";
  str += "<span style=\"color:#0073a5;font-size:small;\">" + user_info.job + "</span>";
  str += "</p>";
  str += "";
  str += "<p>";
  str += "<a style=\"color:#004259 !important';font-size:small;text-decoration:none;\" href=\"https://www.deoever.be\"><strong>www.deoever.be</strong></a>";
  str += "</p>";
  str += "<table style=\"width:500pt;\">";
  str +=   "<colgroup><col style=\"width:10%;\"><col style=\"width:90%;\"></colgroup>";
  str +=   "<tbody>";
  str +=     "<tr>";
  str +=       "<td style=\"height:20px;width:50pt;\">"
  str +=         "<a href=\"https://be.linkedin.com/company/vzw-de-oever\"><img src=\"https://imgmsgen.com/img/bookmark/ln.png\" width=\"20\" height=\"20\"> </a><a href=\"https://www.instagram.com/deoevervzw/\"><img src=\"https://imgmsgen.com/img/bookmark/it.png\" width=\"20\" height=\"20\"></a>";
  str +=       "<td style=\"height:20px;width:450pt;\">";
  str +=         "<span style=\"color:rgb(0,0,0);\"><strong>Partner Agentschap Opgroeien - Jeugdhulp</strong></span><br>"
  str +=         "<span style=\"color:rgb(0,115,165);\">0413.895.634 | RPR Ondernemingsrechtbank Antwerpen afdeling Hasselt</span>"
  str +=       "<td>";
  str +=     "</tr>";
  str +=   "</tbody>";
  str += "</table>";
  str += "<p>";
  str +=   "<img class=\"image_resized\" style=\"width:500pt;\" src=\"https://www.deoever.be/wp-content/uploads/2024/09/banner-algemeen.png\"><br>";
  str += "</p>";
  str += "<p style=\"text-align:justify;color: gray;font-size: xx-small\">";
  str +=   "Deze e-mail en zijn bijlagen zijn uitsluitend bestemd voor de geadresseerde(n) en strikt vertrouwelijk. Hun inhoud kan bij wet beschermd zijn. Indien de mail niet voor u bestemd is, is elke publicatie, reproductie, kopie, distributie of andere verspreiding of gebruik ervan ten strengste verboden. Als u deze boodschap per vergissing toegestuurd kreeg, gelieve de afzender onmiddellijk te verwittigen en de e-mail te vernietigen. Vzw De Oever besteedt de uiterste zorg aan de betrouwbaarheid en actualiteit van de gegevens die het verspreidt. Desalniettemin blijven fouten mogelijk, ook bij de transmissie van de gegevens. De overgebrachte informatie kan onderschept, gewijzigd of vernietigd zijn. Ze kan ook verloren gaan, te laat of onvolledig aankomen of een virus bevatten. Vzw De Oever aanvaardt bijgevolg geen enkele verantwoordelijkheid voor schade als gevolg van onjuistheden of van problemen veroorzaakt door of inherent aan het verspreiden van informatie via e-mail, evenals voor technische storingen en virussen.";
  str += "</p>";

  return str;
}

function get_template_B_str(user_info)
{
  let str = "";
  if (is_valid_data(user_info.greeting))
  {
    str += user_info.greeting + "<br/>";
  }

  str += "<table>";
  str +=   "<tr>";
  str +=     "<td style='border-right: 1px solid #000000; padding-right: 5px;'><img src='https://officedev.github.io/Office-Add-in-samples/Samples/outlook-set-signature/assets/sample-logo.png' alt='Logo' /></td>";
  str +=     "<td style='padding-left: 5px;'>";
  str +=	   "<strong>" + user_info.name + "</strong>";
  str +=     is_valid_data(user_info.pronoun) ? "&nbsp;" + user_info.pronoun : "";
  str +=     "<br/>";
  str +=	   user_info.email + "<br/>";
  str +=	   is_valid_data(user_info.phone) ? user_info.phone + "<br/>" : "";
  str +=     "</td>";
  str +=   "</tr>";
  str += "</table>";

  return str;
}

function get_template_C_str(user_info)
{
  let str = "";
  if (is_valid_data(user_info.greeting))
  {
    str += user_info.greeting + "<br/>";
  }

  str += user_info.name;
  
  return str;
}