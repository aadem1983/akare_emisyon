class TableSort {
    constructor(table) {
        this.table = table;
        // Checkbox sütununu hariç tut
        this.headers = Array.from(table.querySelectorAll('th'))
                          .filter((th, index) => index !== 0); // 0. sütun checkbox
        this.rows = Array.from(table.querySelectorAll('tbody tr'));
        this.sortStates = Array(this.headers.length).fill(0);
        this.init();
    }

    init() {
        // Checkbox sütununu hariç tut
        this.headers.forEach((header, index) => {
            header.addEventListener('click', () => this.sortColumn(index));
        });
    }

    sortColumn(columnIndex) {
        const currentSortState = this.sortStates[columnIndex];
        const newSortState = (currentSortState + 1) % 3;
        this.sortStates.fill(0);
        this.sortStates[columnIndex] = newSortState;

        const sortFunction = (a, b) => {
            // Checkbox sütununu hariç tut
            const cellA = a.cells[columnIndex + 1]; // 0. sütun checkbox olduğu için +1
            const cellB = b.cells[columnIndex + 1];
            const valueA = cellA.textContent.trim();
            const valueB = cellB.textContent.trim();

            if (newSortState === 1) {
                return valueA.localeCompare(valueB);
            } else if (newSortState === 2) {
                return valueB.localeCompare(valueA);
            }
            return 0;
        };

        this.rows.sort(sortFunction);
        this.updateTable();
    }

    updateTable() {
        const tbody = this.table.querySelector('tbody');
        tbody.innerHTML = '';
        this.rows.forEach(row => tbody.appendChild(row));

        // Sıra numaralarını güncelle
        this.rows.forEach((row, index) => {
            row.cells[1].textContent = index + 1;
        });
    }
}
