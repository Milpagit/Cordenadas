async function getCoordinates(address, fallbackLocation) {
  const apiKey = "AIzaSyAQw94hlDpP-gOD_2ZeVDxTptg5bUUUs7E"; // Reemplaza con tu clave de Google Maps
  const url = `https://maps.googleapis.com/maps/api/geocode/json?address=${encodeURIComponent(
    address
  )}&key=${apiKey}`;

  try {
    const response = await fetch(url);
    const data = await response.json();

    if (data.status === "OK") {
      return data.results[0].geometry.location;
    } else {
      console.log(`Dirección no encontrada: ${address}`);

      // Si no se encuentra, intentamos con la localidad
      if (fallbackLocation) {
        console.log(`Intentando con la localidad: ${fallbackLocation}`);
        return getCoordinates(fallbackLocation, null);
      } else {
        return { lat: "N/A", lng: "N/A" };
      }
    }
  } catch (error) {
    console.error("Error en la API:", error);
    return { lat: "Error", lng: "Error" };
  }
}

async function processExcel() {
  const fileInput = document.getElementById("fileInput");
  const tableBody = document.querySelector("#resultTable tbody");

  if (fileInput.files.length === 0) {
    alert("Por favor, selecciona un archivo Excel.");
    return;
  }

  const file = fileInput.files[0];
  const reader = new FileReader();

  reader.onload = async function (e) {
    const data = new Uint8Array(e.target.result);
    const workbook = XLSX.read(data, { type: "array" });
    const sheet = workbook.Sheets[workbook.SheetNames[0]];
    const jsonData = XLSX.utils.sheet_to_json(sheet);

    tableBody.innerHTML = ""; // Limpiar tabla antes de cargar nuevos datos

    for (let row of jsonData) {
      const address = row["Dirección"];
      const locality = row["Localidad"]; // Nueva columna

      if (!address && !locality) continue; // Si no hay datos, pasamos

      const coordinates = await getCoordinates(address, locality);
      const rowHtml = `<tr>
                                <td class="no-select">${
                                  address || locality
                                }</td>
                                <td>${coordinates.lat}</td>
                                <td>${coordinates.lng}</td>
                             </tr>`;
      tableBody.innerHTML += rowHtml;
    }
  };

  reader.readAsArrayBuffer(file);
}
