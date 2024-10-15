document.addEventListener('DOMContentLoaded', function() {
    if (document.getElementById('chart')) {
        drawChart(filteredData);
    }

    const filterForm = document.getElementById('filter-form');
    if (filterForm) {
        const filters = filterForm.querySelectorAll('select');
        filters.forEach(filter => {
            filter.addEventListener('change', () => {
                copyFiltersToHiddenInputs();
                filterForm.submit();
            });
        });
    }

    const multiSelectContainer = document.querySelector('.multi-select-container');
    const multiSelectSelected = multiSelectContainer.querySelector('.multi-select-selected');
    const multiSelectDropdown = multiSelectContainer.querySelector('.multi-select-dropdown');
    const multiSelectOptions = multiSelectDropdown.querySelectorAll('.multi-select-option input[type="checkbox"]');
    const selectAllCheckbox = multiSelectDropdown.querySelector('#select-all');
    const clearAllCheckbox = multiSelectDropdown.querySelector('#clear-all');
    const applyButton = multiSelectDropdown.querySelector('.multi-select-apply');

    // Toggle dropdown visibility on click
    multiSelectSelected.addEventListener('click', function (event) {
        event.stopPropagation();  // Prevent the event from bubbling up
        multiSelectDropdown.classList.toggle('show');
    });

    // Handle select all functionality
    selectAllCheckbox.addEventListener('click', function () {
        multiSelectOptions.forEach(option => {
            option.checked = true;  // Corrected boolean value
        });
        updateSelectedItems();
    });

    clearAllCheckbox.addEventListener('click', function () {
        // Clear all checkboxes when clicked
        multiSelectOptions.forEach(option => {
            option.checked = false;
        });
        selectAllCheckbox.checked = false;
        updateSelectedItems();
    
        // Reset the clear all checkbox state immediately
        clearAllCheckbox.checked = false;
    });
    

    // Apply button functionality
    applyButton.addEventListener('click', function (event) {
        event.preventDefault();  // Prevent the form from submitting

        copyFiltersToHiddenInputs();  // Capture and log the selected countries

        console.log('Submitting form with selected countries:', document.getElementById('hidden-country').value);  // Log before submission

        // Manually submit the form if everything looks correct
        filterForm.submit();
    });

    // Update the selected items display when an option is changed
    multiSelectOptions.forEach(option => {
        option.addEventListener('change', function () {
            updateSelectedItems();
        });
    });

    function updateSelectedItems() {
        const selectedItems = Array.from(multiSelectOptions)
            .filter(option => option.checked)
            .map(option => option.nextElementSibling.textContent);
        multiSelectSelected.innerHTML = selectedItems.join(', ') || 'Select Country';
    }

    // Close the dropdown when clicking outside of it
    document.addEventListener('click', function (e) {
        if (!multiSelectContainer.contains(e.target) && !multiSelectSelected.contains(e.target)) {
            multiSelectDropdown.classList.remove('show');
        }
    });

    // Restore the selected values for the country multiselect
    const hiddenCountryInput = document.getElementById('hidden-country');
    const selectedCountryValues = hiddenCountryInput.value.split(',');
    console.log("Restoring Countries: ", selectedCountryValues);

    selectedCountryValues.forEach(value => {
        const checkbox = document.querySelector(`.multi-select-container .multi-select-option input[type="checkbox"][value="${value}"]`);
        if (checkbox) {
            checkbox.checked = true;
        }
    });

    // Update the displayed selected items
    updateSelectedItems();
});


function copyFiltersToHiddenInputs() {
    // Copy other filters to hidden inputs
    // document.getElementById('hidden-relationship').value = document.getElementById('relationship').value;
    document.getElementById('hidden-industry').value = document.getElementById('industry').value;
    document.getElementById('hidden-revenue').value = document.getElementById('revenue').value;
    document.getElementById('hidden-industry_primary').value = document.getElementById('industry_primary').value
    document.getElementById('hidden-combined-opp-score').value = document.getElementById('combined_opp_score').value
    document.getElementById('hidden-priority-status').value = document.getElementById('priority_status').value
    document.getElementById('hidden-country').value = document.getElementById('country').value;
    document.getElementById('hidden-geography').value = document.getElementById('geography').value;



    

    // Collect all selected countries and ensure they are passed as an array
    const selectedCountries = Array.from(document.querySelectorAll('.multi-select-container .multi-select-option input[type="checkbox"]:checked'))
        .map(checkbox => checkbox.value);

    console.log("Selected Countries: ", selectedCountries);  // Log selected countries

    // Update the hidden input for country[]
    document.getElementById('hidden-country').value = selectedCountries.join('|');
}


function drawChart(chartData) {
    console.log("Chart Data: ", chartData);

    const ctx = document.getElementById('chart').getContext('2d');
    let minRevenue = Math.min(...chartData.map(d => parseFloat(d['Revenue (M)'])));
    let maxRevenue = Math.max(...chartData.map(d => parseFloat(d['Revenue (M)'])));
    

    chartData = chartData.map(d => ({
        x: d['Combined Opp. Score'] || 0,
        y: d['Bain Relationship Score'] || 0,
        r: d['Revenue (M)'] < 5000 ? parseFloat((d['Revenue (M)'] / 100).toFixed(0)) :((parseFloat(d['Revenue (M)']) - minRevenue) / (maxRevenue - minRevenue) * 100).toFixed(0),
        //r: parseFloat(d['Revenue (M)']/10000).toFixed(0),
        //r: d['Revenue (M)'] < 5000 ? parseFloat((d['Revenue (M)'] / 100).toFixed(0)) : parseFloat((d['Revenue (M)'] / 10000).toFixed(0)) || 0,
        company: d['Company'] || '',
        Rev: parseFloat((d['Revenue (M)'] / 1000).toFixed(1)) || 0
    }));

    chartData.sort((a, b) => a.r - b.r);
    console.log('Sorted Chart Data:', chartData);

    const xValues = chartData.map(d => d.x);
    const yValues = chartData.map(d => d.y);

    const minX = 0;
    const maxX = 25;
    const minY = 0;
    const maxY = 12;
    console.log('Axis values', minX, maxX, minY, maxY);

    const data = {
        datasets: [{
            data: chartData,
            backgroundColor: chartData.map(d => priorityCompanies.includes(d.company) ? 'rgba(204, 0, 0, 0.75)' : 'rgba(92, 92, 92, 0.75)'),
            borderColor: chartData.map(d => priorityCompanies.includes(d.company) ? 'rgba(204, 0, 0, 0.75)' : 'rgba(92, 92, 92, 0.75)'),
            borderWidth: 1
        }]
    };

    const config = {
        type: 'bubble',
        data: data,
        options: {
            maintainAspectRatio: false,
            scales: {
                x: {
                    ticks: {
                        color: '#000000'
                    },
                    min: minX,
                    max: maxX,
                    grid: {
                        display: false
                    }
                },
                y: {
                    ticks: {
                        color: '#000000'
                    },
                    min: minY,
                    max: maxY,
                    grid: {
                        display: false
                    }
                }
            },
            plugins: {
                tooltip: {
                    callbacks: {
                        label: function(context) {
                            let label = `Company: ${context.raw.company}`;
                            label += `, Revenue: $${context.raw.Rev}B`;
                            return label;
                        }
                    }
                },
                legend: {
                    display: false,
                    labels: {
                        color: '#000000',
                        font: {
                            size: 12
                        }
                    }
                },
                datalabels: {
                    align: 'center',
                    anchor: 'center',
                    color: 'white',
                    font: {
                        weight: 'bold',
                        size: 16
                    },
                    formatter: function(value, context) {
                        return value.company;
                    }
                }
            }
        }
    };

    const myChart = new Chart(ctx, config);
    myChart.update();
}
