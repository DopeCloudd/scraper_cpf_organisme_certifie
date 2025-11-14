import axios from "axios";

async function scrapeSite() {
  // perform an HTTP GET request to the target page
  const response = await axios.get(
    "https://www.moncompteformation.gouv.fr/espace-prive/html/#/formation/recherche/91881917800016_BC24DIJON/91881917800016_BC24DIJON?contexteFormation=ACTIVITE_PROFESSIONNELLE"
  );
  // get the HTML from the server response
  // and log it
  const html = response.data;
  console.log(html);
}

scrapeSite();
