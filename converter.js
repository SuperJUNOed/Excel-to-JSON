let blob;

function readExcelFile(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.readAsBinaryString(file);

    reader.onload = (e) => {
      const data = e.target.result;
      const workbook = XLSX.read(data, { type: "binary" });
      //Get all sheet names
      const sheetNames = workbook.SheetNames;
      //Create an object to hold all worksheet data
      let allData = {};
      let newLabel;
      let descrip;
      //Loop through each sheet
      sheetNames.forEach((sheetName) => {
        //Get the worksheet for the current sheet
        const worksheet = workbook.Sheets[sheetName];
        // Convert the worksheet to JSON
        const excelData = XLSX.utils.sheet_to_json(worksheet);
        let des = description(excelData);
        let act = actors(excelData);
        let pur = purposes(excelData);
        let person = personCategories(excelData);
        let datacategory = dataCategories(excelData);
        let rec = recipients(excelData);
        let secu = securityMeasures(excelData);
        const newDes = { tags: [] };
        newValue = des["description"][0].value;
        if (newValue) {
          if (newValue.includes(",")) {
            const commaLabel = newValue.split(", ");
            commaLabel.map((item) => {
              const newData = item.trim();
              newDes.tags.push({
                label: newData,
              });
            });
          } else {
            newDes.tags.push({
              label: des["description"][0].value,
            });
          }
        }
        if (des["description"][1].value.length >= 120) {
          newLabel = des["description"][1].value.substring(0, 119);
          descrip = des["description"][1].value.substring(
            120,
            des["description"][1].value.length - 1
          );
        } else if (des["description"][1].value.length < 120) {
          newLabel = des["description"][1].value;
          descrip = " ";
        } else {
          newLabel = " ";
          descrip = " ";
        }
        let oneData = {
          ref: sheetName,
          label: newLabel,
          description: descrip,
          dateCreation: new Date().toISOString(),
          dateUpdate: new Date().toISOString(),
          processingType: "Default",
          processingState: "InProduction",
        };
        oneData = {
          ...oneData,
          ...newDes,
          ...act,
          ...pur,
          ...person,
          ...datacategory,
          ...rec,
          ...secu,
        };

        //Add the JSON data to the allData object
        allData[sheetName] = oneData;
        resolve(allData);
      });
    };

    reader.onerror = (e) => {
      reject(e);
    };
  });
}

function convertExcelToJson(input) {
  // Display uploaded file
  readURL(input);
  // Get data from file
  readExcelFile(input.files[0]).then((data) => {
    // Format data
    const jData = {
      $schema: "https://api.dastra.eu/v1/jsonschema/data-processing-record",
      version: "1.2",
      date: new Date().toISOString(),
      label: "Export : " + new Date().toLocaleDateString(),
      records: Object.values(data),
    };
    // Save the JSON data to a file
    const jsonData = JSON.stringify(jData, null, 2);
    blob = new Blob([jsonData], { type: "application/json" });
    saveAs(blob, "data.json");
  });
}

function downloadJSON() {
  saveAs(blob, "data.json");
}

function description(data) {
  const newData = data.slice(1, 5);
  const description = newData.map((item) => ({
    label: item["__EMPTY"],
    value: item["Fiche de registre"],
  }));
  return { description: description };
}

function actors(data) {
  const actors = { assets: [] };

  const subStr1 = "jusqu'en";
  const subStr2 = "depuis";

  // loop through data slice
  data.slice(12, 13).forEach((item) => {
    if (!item["Fiche de registre"]) return; // skip empty items
    if (item["Fiche de registre"].includes(",")) {
      const commaLabel = item["Fiche de registre"].split(",");
      commaLabel.map((item) => {
        const newData = item.trim();
        actors.assets.push({
          label: newData,
          type: "Software",
        });
      });
    }
    // check if label contains "-" separator
    if (item["Fiche de registre"].includes("-")) {
      const newLabel = item["Fiche de registre"].split("-");
      let index1 = newLabel[0].toLowerCase().indexOf(subStr1);
      let result1 = newLabel[0]
        .toLowerCase()
        .slice(index1, index1 + subStr1.length);
      let newStr1 = newLabel[0]
        .substring(0, newLabel[0].toLowerCase().indexOf(result1))
        .trim();
      // const label1 = newStr1 ? newStr1 : "";

      let index2 = newLabel[1].toLowerCase().indexOf(subStr2);
      let result2 = newLabel[1]
        .toLowerCase()
        .slice(index2, index2 + subStr2.length);
      let newStr2 = newLabel[1]
        .substring(0, newLabel[1].toLowerCase().indexOf(result2))
        .trim();
      // const label2 = newStr2 ? newStr2 : "";

      // add extracted labels to assets array
      actors.assets.push({ label: newStr1 });
      actors.assets.push({ label: newStr2 });
    }
  });

  return actors;
}

function purposes(data) {
  let purposes = {};
  let newPurposes = {};
  let newData = {};
  newData = data.slice(14, 22);
  newData.map((item, index) => {
    purposes["purposes"] = purposes["purposes"] || [];
    if (item["Fiche de registre"]) {
      purposes["purposes"].push({
        purposes: item["__EMPTY"],
        "purposes.legalbasis": item["Fiche de registre"],
      });
    }
  });
  const legalbasis = purposes["purposes"].find(
    (obj) => obj.purposes == "Base légale : Consentement"
  );
  if (legalbasis == undefined) {
    purposes["purposes"].push({
      purposes: "Base légale : Consentement",
      "purposes.legalbasis": null,
    });
  }
  for (var i = 0; i < purposes["purposes"].length; i++) {
    newPurposes["purposes"] = newPurposes["purposes"] || [];
    if (!purposes["purposes"][i]["purposes"].startsWith("Base")) {
      if (purposes["purposes"][i]["purposes.legalbasis"].length > 120) {
        newPurposes["purposes"].push({
          label: purposes["purposes"][i]["purposes.legalbasis"].substring(
            0,
            119
          ),
          description: purposes["purposes"][i]["purposes.legalbasis"].substring(
            120,
            purposes["purposes"][i]["purposes.legalbasis"].length - 1
          ),
          legalBasis:
            purposes["purposes"][purposes["purposes"].length - 1][
              "purposes.legalbasis"
            ],
        });
      } else {
        newPurposes["purposes"].push({
          label: purposes["purposes"][i]["purposes.legalbasis"],
          legalBasis:
            purposes["purposes"][purposes["purposes"].length - 1][
              "purposes.legalbasis"
            ],
        });
      }
    }
  }
  if (newPurposes["purposes"]) {
    for (var i = 0; i < newPurposes["purposes"].length; i++) {
      if (newPurposes["purposes"][i]["legalBasis"] == "Intérêt Légitime") {
        newPurposes["purposes"][i]["legalBasis"] = "LegitimateInterest";
      }
      if (newPurposes["purposes"][i]["legalBasis"] == "Contrat") {
        newPurposes["purposes"][i]["legalBasis"] = "Contract";
      }
      if (newPurposes["purposes"][i]["legalBasis"] == "Consentement") {
        newPurposes["purposes"][i]["legalBasis"] = "Consent";
      }
      if (newPurposes["purposes"][i]["legalBasis"] == "Obligations Légales") {
        newPurposes["purposes"][i]["legalBasis"] = "LegalCommitment";
      }
      if (newPurposes["purposes"][i]["legalBasis"] == "A définir") {
        newPurposes["purposes"][i]["legalBasis"] = null;
      }
    }
  }
  return newPurposes;
}

function dataCategories(data) {
  let des = description(data);
  const newData = data.slice(28, 47);
  let dataFields = { DataFields: [] };
  newData.map((item) => {
    if (item["Fiche de registre"] && item["__EMPTY"] != "Données sensibles") {
      let newItem = item["Fiche de registre"].split(",");
      newItem.map((items) => {
        if (items && items != " ") {
          if (items.length > 150) {
            dataFields.DataFields.push({
              label: items.substring(0, 149).trim(),
              sensitiveData: false,
              personalDataCategory: item["__EMPTY"],
            });
          } else {
            dataFields.DataFields.push({
              label: items.trim(),
              sensitiveData: false,
              personalDataCategory: item["__EMPTY"],
            });
          }
        } else return;
      });
    }
    dataFields.DataFields.map((item) => {
      if (
        item["personalDataCategory"] ==
        "Etat civil, identité, données d'identification, images…"
      ) {
        item["personalDataCategory"] = "CivilStatus";
      }
      if (
        item["personalDataCategory"] ==
        "Vie personnelle (habitudes de vie, situation familiale, ...)"
      ) {
        item["personalDataCategory"] = "PersonalLife";
      }
      if (
        item["personalDataCategory"] ==
        "Vie professionnelle (CV, situation professionnelles, scolarité, formation, distinctions, diplômes, ...)"
      ) {
        item["personalDataCategory"] = "ProfessionalLife";
      }
      if (
        item["personalDataCategory"] ==
        "Informations d'ordre économique et financier (revenus, situation financière, situation fiscale, données bancaires ...)"
      ) {
        item["personalDataCategory"] = "EconomicFinancialData";
      }
      if (
        item["personalDataCategory"] ==
        "Données de connexion (adress IP, logs, identifiants des terminaux, identifiants de connexion, informations d'horodatage ...)"
      ) {
        item["personalDataCategory"] = "ConnectionData";
      }
      if (
        item["personalDataCategory"] ==
        "Données de localisation (déplacements, données GPS, GSM, ...)"
      ) {
        item["personalDataCategory"] = "GeoLocationData";
      }
      if (
        item["personalDataCategory"] ==
        "Internet (cookies, traceurs, données de navigation, mesures d’audience, …)"
      ) {
        item["personalDataCategory"] = "InternetData";
      }
      if (
        item["personalDataCategory"] ==
        "Autres catégories de données (précisez) : "
      ) {
        item["personalDataCategory"] = "Other";
      }
      if (
        item["personalDataCategory"] ==
        "Données révèlant l'origine raciale ou ethnique"
      ) {
        item["personalDataCategory"] = "EthnicalData";
      }
      if (
        item["personalDataCategory"] ==
        "Données révèlant les opinions politiques"
      ) {
        item["personalDataCategory"] = "PoliticalOpinions";
      }
      if (
        item["personalDataCategory"] ==
        "Données révèlant les convictions religieuses ou philosophiques "
      ) {
        item["personalDataCategory"] = "ReligiousBeliefs";
      }
      if (
        item["personalDataCategory"] ==
        "Données révèlant l'appartenance syndicale"
      ) {
        item["personalDataCategory"] = "TradeUnionMembership";
      }
      if (item["personalDataCategory"] == "Données génétiques") {
        item["personalDataCategory"] = "GeneticData";
      }
      if (
        item["personalDataCategory"] ==
        "Données biométriques aux fins d'identifier une personne physique de manière unique"
      ) {
        item["personalDataCategory"] = "BiometricData";
      }
      if (item["personalDataCategory"] == "Données concernant la santé") {
        item["personalDataCategory"] = "HealthData";
      }
      if (
        item["personalDataCategory"] ==
        "Données concernant la vie sexuelle ou l'orientation sexuelle "
      ) {
        item["personalDataCategory"] = "SexualOrientations";
      }
      if (
        item["personalDataCategory"] ==
        "Données relatives à des condamnations pénales ou  infractions"
      ) {
        item["personalDataCategory"] = "CriminalConvictions";
      }
      if (
        item["personalDataCategory"] ==
        "Numéro d'identification  national unique (NIR pour la France)"
      ) {
        item["personalDataCategory"] = "NIR";
      }
    });
  });
  if (dataFields.DataFields.length > 0)
    dataFields.DataFields[
      dataFields.DataFields.length - 1
    ].sensitiveData = true;
  const dataCategories = {
    label: "Données du traitement" + " " + des["description"][1].value,
    dataFields: [...dataFields.DataFields],
  };
  return { dataRetentionRules: [dataCategories] };
}

function personCategories(data) {
  let personCategories = { personCategories: [] };
  const newData = data.slice(23, 27);
  newData.map((item) => {
    personCategories["personCategories"] =
      personCategories["personCategories"] || [];
    if (item["Fiche de registre"]) {
      if (item["Fiche de registre"].includes(",")) {
        const commaLabel = item["Fiche de registre"].split(",");
        commaLabel.map((items) => {
          if (items && items != " ") {
            const newData = items.trim();
            personCategories["personCategories"].push({
              subjectCategory: { label: newData },
            });
          }
        });
      } else {
        personCategories["personCategories"].push({
          subjectCategory: {
            label: item["Fiche de registre"],
          },
        });
      }
    }
  });

  return personCategories;
}

function recipients(data) {
  let recipients = {};
  let newData = {};
  newData = data.slice(48, 51);
  newData.map((item, index) => {
    recipients["recipients"] = recipients["recipients"] || [];
    if (item["Fiche de registre"] && item["Fiche de registre"].includes(",")) {
      const commaLabel = item["Fiche de registre"].split(",");
      if (
        item["Fiche de registre"] &&
        item["__EMPTY"] !=
          "Sous-traitants (Exemples : hébergeurs, prestataires et maintenance informatiques, ...)"
      ) {
        commaLabel.map((items) => {
          if (items) {
            const newData = items.trim();
            recipients["recipients"].push({
              recipientType: item["__EMPTY"],
              label: newData,
            });
          } else return;
        });
      }
      if (
        item["__EMPTY"] ==
        "Sous-traitants (Exemples : hébergeurs, prestataires et maintenance informatiques, ...)"
      ) {
        commaLabel.map((items) => {
          if (items) {
            const newData = items.trim();
            recipients["recipients"].push({
              recipientType: item["__EMPTY"],
              actor: { companyName: newData },
            });
          } else return;
        });
      }
    } else {
      if (
        item["Fiche de registre"] &&
        item["__EMPTY"] !=
          "Sous-traitants (Exemples : hébergeurs, prestataires et maintenance informatiques, ...)"
      ) {
        if (item["Fiche de registre"].length > 250) {
          recipients["recipients"].push({
            recipientType: item["__EMPTY"],
            label: item["Fiche de registre"].substring(0, 249),
          });
        } else {
          recipients["recipients"].push({
            recipientType: item["__EMPTY"],
            label: item["Fiche de registre"],
          });
        }
      }
      if (
        item["__EMPTY"] ==
        "Sous-traitants (Exemples : hébergeurs, prestataires et maintenance informatiques, ...)"
      ) {
        if (item["Fiche de registre"]) {
          recipients["recipients"].push({
            recipientType: item["__EMPTY"],
            actor: { companyName: item["Fiche de registre"] },
          });
        } else return;
      }
    }
  });
  for (var i = 0; i < recipients["recipients"].length; i++) {
    if (
      recipients["recipients"][i].recipientType ==
      "Destinataires Internes (exemples : entité ou service, catégories de personnes habilitées ...)"
    ) {
      recipients["recipients"][i].recipientType = "InternalService";
    }
    if (
      recipients["recipients"][i].recipientType ==
      "Organismes externes (Exemples : filiales, partenaires, ...)"
    ) {
      recipients["recipients"][i].recipientType = "ThirdParty";
    }
    if (
      recipients["recipients"][i].recipientType ==
      "Sous-traitants (Exemples : hébergeurs, prestataires et maintenance informatiques, ...)"
    ) {
      recipients["recipients"][i].recipientType = "Vendor";
    }
  }
  return recipients;
}

function securityMeasures(data) {
  const securityMeasures = { securityMeasures: [] };

  for (const item of data.slice(53, 55)) {
    if (item["__EMPTY_1"]) {
      const replaceCross = item["__EMPTY_1"].replace(/[*]/g, "");
      const valueCross = replaceCross
        .split(/\r?\n/)
        .map((item, index) => !!item && index)
        .filter((e) => e);
      const replaceData = item["Fiche de registre"].replace(/[*]/g, "");
      const valueData = replaceData.split(/\r?\n/);
      const correctData = valueCross.map((i) => valueData[i]);
      const flag = item["__EMPTY"].endsWith("techniques")
        ? "Technical"
        : "Organizational";

      for (const value of correctData) {
        if (value) {
          securityMeasures["securityMeasures"].push({
            label: value,
            type: flag,
          });
        }
      }
    } else return;
  }

  return securityMeasures;
}
