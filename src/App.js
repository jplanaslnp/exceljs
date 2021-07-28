import './App.css';
import ExcelJS from "exceljs";
import saveAs from 'file-saver';

function App() {

  const submitForm = e => {
    e.preventDefault();
  
    const wb = new ExcelJS.Workbook();
    const reader = new FileReader();
    var file = document.getElementById("testFile");
    reader.readAsArrayBuffer(file.files[0]);
    reader.onload = () => {
      const buffer = reader.result;
      wb.xlsx.load(buffer).then(workbook => {
        console.log(workbook, "workbook instance");
        workbook.eachSheet((sheet, id) => {
          sheet.eachRow((row, rowIndex) => {
            console.log(row.values, rowIndex);
          });
        });
      });
    };
  };


  const createExcel = () => {
    console.log("create excel")
    const workbook = new ExcelJS.Workbook();
    const worksheet =  workbook.addWorksheet('Mysheet');


    worksheet.columns = [
      {header: 'Id', key: 'id', width: 10},
      {header: 'Name', key: 'name', width: 32}, 
      {header: 'D.O.B.', key: 'dob', width: 15,},
      {header: 'A1', key: 'A1', width: 15,}
     ];

     worksheet.addRow({id: 1, name: 'John Doe', dob: new Date(1970, 1, 1), A1:1});
     worksheet.addRow({id: 2, name: 'Jane Doe', dob: new Date(1965, 1, 7), A1:2});

      const row = worksheet.getRow(1);
      console.log(row.getCell(1))
      console.log(row.getCell(2))
      console.log(worksheet.getCell('A1').value)
      

      row.getCell(1).dataValidation = {
        type: 'list',
        allowBlank: true,
        formulae: ['"One,Two,Three,Four"']
      }

      console.log("sfdsdfd ", worksheet.getCell('A1'))

    //await workbook.xlsx.writeFile("hola");
    workbook.xlsx.writeBuffer().then(buffer => saveAs(new Blob([buffer]), `${Date.now()}_feedback.xlsx`)).catch(err => console.log('Error writing excel export', err))
  };



  return (
    <div className="App">
      <div>
        <form onSubmit={submitForm} id="testForm">
          <input type="text" id="testName" />
          <input type="file" name="test" id="testFile" />
          <button type="submit">Submit</button>
        </form>
      </div>
      <button onClick={createExcel}>Submit</button>
    </div>
  );
}

export default App;
