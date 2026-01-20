Office.onReady(function() {
    // Office is ready
});

async function copyCellForLLM(event) {
    try {
        await Excel.run(async (context) => {
            const range = context.workbook.getSelectedRange();
            range.load(["address", "values", "rowCount", "columnCount"]);

            await context.sync();

            // Parse the address to get just the cell references (without sheet name)
            const fullAddress = range.address;
            const cellsAddress = fullAddress.includes("!")
                ? fullAddress.split("!")[1]
                : fullAddress;

            let output = "";

            if (range.rowCount === 1 && range.columnCount === 1) {
                // Single cell
                const value = range.values[0][0];
                output = `${cellsAddress}: ${formatValue(value)}`;
            } else {
                // Multiple cells - iterate through the range
                const lines = [];
                const startAddress = cellsAddress.includes(":")
                    ? cellsAddress.split(":")[0]
                    : cellsAddress;

                const startCol = startAddress.match(/[A-Z]+/)[0];
                const startRow = parseInt(startAddress.match(/[0-9]+/)[0]);

                for (let row = 0; row < range.rowCount; row++) {
                    for (let col = 0; col < range.columnCount; col++) {
                        const cellRef = columnToLetter(letterToColumn(startCol) + col) + (startRow + row);
                        const value = range.values[row][col];
                        lines.push(`${cellRef}: ${formatValue(value)}`);
                    }
                }
                output = lines.join("\n");
            }

            // Copy to clipboard
            await navigator.clipboard.writeText(output);
        });
    } catch (error) {
        console.error("Error: " + error);
    }

    event.completed();
}

function formatValue(value) {
    if (value === null || value === undefined || value === "") {
        return "[empty]";
    }
    return String(value);
}

function letterToColumn(letter) {
    let column = 0;
    for (let i = 0; i < letter.length; i++) {
        column = column * 26 + (letter.charCodeAt(i) - 64);
    }
    return column;
}

function columnToLetter(column) {
    let letter = "";
    while (column > 0) {
        let temp = (column - 1) % 26;
        letter = String.fromCharCode(temp + 65) + letter;
        column = Math.floor((column - temp - 1) / 26);
    }
    return letter;
}

Office.actions.associate("copyCellForLLM", copyCellForLLM);
