document.addEventListener('DOMContentLoaded', function () {
    const addRowBtn = document.getElementById('addRowBtn');
    const valueTable = document.getElementById('valueTable').getElementsByTagName('tbody')[0];

    addRowBtn.addEventListener('click', function () {
        const newRow = document.createElement('tr');
        newRow.innerHTML = `
            <td><input type="text" name="userId[]" required></td>
            <td><input type="text" name="userName[]" required></td>
            <td><input type="text" name="segment[]" required></td>
            <td><input type="text" name="approvedAddress[]" required></td>
            <td><input type="text" name="actualLocatedAt[]" required></td>
            <td><input type="text" name="remarks[]" required></td>
            <td><input type="text" name="authorizedPersonReply[]" required></td>
            <td><input type="file" name="evidence[]" accept="image/*"></td>
            <td><button type="button" class="remove-row-btn">Remove</button></td>
        `;
        valueTable.appendChild(newRow);
    });

    valueTable.addEventListener('click', function (event) {
        if (event.target.classList.contains('remove-row-btn')) {
            event.target.parentElement.parentElement.remove();
        }
    });
});
