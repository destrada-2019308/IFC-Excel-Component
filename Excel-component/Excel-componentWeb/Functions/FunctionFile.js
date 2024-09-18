// La función initialize debe ejecutarse cada vez que se carga una página nueva.
Office.onReady(() => {
        // Si necesita inicializar algo, puede hacerlo aquí.
});

async function sampleFunction(event) { 
const values = [
        [Math.floor(Math.random() * 1000), Math.floor(Math.random() * 1000), Math.floor(Math.random() * 1000)],
        [Math.floor(Math.random() * 1000), Math.floor(Math.random() * 1000), Math.floor(Math.random() * 1000)],
        [Math.floor(Math.random() * 1000), Math.floor(Math.random() * 1000), Math.floor(Math.random() * 1000)]
        ];

        try {
        await Excel.run(async (context) => {
                // Write sample values to a range in the active worksheet.
                const sheet = context.workbook.worksheets.getActiveWorksheet();
                sheet.getRange("B3:D5").values = values;
                await context.sync();
        });
        } catch (error) {
        console.log(error.message);
        }
        // Es necesario llamar a event.completed. event.completed permite a la plataforma saber que se ha completado el procesamiento.
        event.completed();
}
