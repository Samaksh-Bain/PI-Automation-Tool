document.addEventListener('DOMContentLoaded', function () {
    // Check if the chart element exists, then draw the chart using filtered data
    if (document.getElementById('chart')) {
        drawChart(filteredData);
    }

    // Get the filter form element
    const filterForm = document.getElementById('filter-form');
    if (filterForm) {
        const filters = filterForm.querySelectorAll('select');
        filters.forEach(filter => {
            filter.addEventListener('change', () => {
                copyFiltersToHiddenInputs(); // Copy filter values to hidden inputs
                filterForm.submit(); // Submit the form with updated filters
            });
        });
    }


    // -------------------- Country Filter Logic --------------------
    const countryContainer = document.querySelector('#country-filter');
    const countryDropdown = countryContainer.querySelector('.multi-select-dropdown');
    const countrySelected = countryContainer.querySelector('#country');
    const countryOptions = countryDropdown.querySelectorAll('.multi-select-option input[type="checkbox"]');
    const selectAllCountryButton = countryDropdown.querySelector('#select-all');
    const clearAllCountryButton = countryDropdown.querySelector('#clear-all');
    const applyCountryButton = countryDropdown.querySelector('.multi-select-apply');
    const searchCountryInput = countryDropdown.querySelector('.multi-select-search input[type="text"]');

    // Toggle dropdown visibility when the selected area is clicked (Country)
    countrySelected.addEventListener('click', function (event) {
        event.stopPropagation();  // Prevent the event from bubbling up
        countryDropdown.classList.toggle('show');
        console.log('Toggling Country dropdown');
    });

    // Handle select all functionality for Country
    selectAllCountryButton.addEventListener('click', function () {
        countryOptions.forEach(option => {
            option.checked = true;
        });
        updateSelectedCountryItems();
        filterForm.submit();
    });

    // Handle "Clear All" functionality for Country
    clearAllCountryButton.addEventListener('click', function () {
        countryOptions.forEach(option => {
            option.checked = false;
        });
        updateSelectedCountryItems();
    });

    // Apply button functionality for Country
    applyCountryButton.addEventListener('click', function (event) {
        event.preventDefault();
        copyCountryFiltersToHiddenInputs();
        countryDropdown.classList.remove('show');
        updateSelectedCountryItems();
        filterForm.submit();
    });

    // Search functionality for Country
    searchCountryInput.addEventListener('input', function () {
        const filter = this.value.toLowerCase();
        countryOptions.forEach(option => {
            const label = option.nextElementSibling.textContent.toLowerCase();
            option.closest('.multi-select-option').style.display = label.includes(filter) ? '' : 'none';
        });
    });

    // Update selected items display for Country
    function updateSelectedCountryItems() {
        const selectedItems = Array.from(countryOptions)
            .filter(option => option.checked)
            .map(option => option.nextElementSibling.textContent);
        countrySelected.innerHTML = selectedItems.join(', ') || 'Select Country';
    }

    // Copy selected countries to hidden input
    function copyCountryFiltersToHiddenInputs() {
        const selectedCountries = Array.from(countryOptions)
            .filter(option => option.checked)
            .map(option => option.value);
        document.getElementById('hidden-country').value = selectedCountries.join('|');
    }

    // -------------------- Company Filter Logic --------------------
    const companyContainer = document.querySelector('#company-filter');
    const companyDropdown = companyContainer.querySelector('.multi-select-dropdown');
    const companySelected = companyContainer.querySelector('#company');
    const companyOptions = companyDropdown.querySelectorAll('.multi-select-option input[type="checkbox"]');
    const selectAllCompanyButton = companyDropdown.querySelector('#select-all-company');
    const clearAllCompanyButton = companyDropdown.querySelector('#clear-all-company');
    const applyCompanyButton = companyDropdown.querySelector('.multi-select-apply');
    const searchCompanyInput = companyDropdown.querySelector('.multi-select-search input[type="text"]');

    // Toggle dropdown visibility when the selected area is clicked (Company)
    companySelected.addEventListener('click', function (event) {
        event.stopPropagation();  // Prevent the event from bubbling up
        companyDropdown.classList.toggle('show');
        console.log('Toggling Company dropdown');
    });

    // Handle select all functionality for Company
    selectAllCompanyButton.addEventListener('click', function () {
        companyOptions.forEach(option => {
            option.checked = true;
        });
        updateSelectedCompanyItems();
        filterForm.submit();
    });

    // Handle "Clear All" functionality for Company
    clearAllCompanyButton.addEventListener('click', function () {
        companyOptions.forEach(option => {
            option.checked = false;
        });
        updateSelectedCompanyItems();
        filterForm.submit();
    });

    // Apply button functionality for Company
    applyCompanyButton.addEventListener('click', function (event) {
        event.preventDefault();
        copyCompanyFiltersToHiddenInputs();
        companyDropdown.classList.remove('show');
        updateSelectedCompanyItems();
        filterForm.submit();
    });

    // Search functionality for Company
    searchCompanyInput.addEventListener('input', function () {
        const filter = this.value.toLowerCase();
        companyOptions.forEach(option => {
            const label = option.nextElementSibling.textContent.toLowerCase();
            option.closest('.multi-select-option').style.display = label.includes(filter) ? '' : 'none';
        });
    });

    // Update selected items display for Company
    function updateSelectedCompanyItems() {
        const selectedItems = Array.from(companyOptions)
            .filter(option => option.checked)
            .map(option => option.nextElementSibling.textContent);
        companySelected.innerHTML = selectedItems.join(', ') || 'Select Company';
    }

    // Copy selected companies to hidden input
    function copyCompanyFiltersToHiddenInputs() {
        const selectedCompanies = Array.from(companyOptions)
            .filter(option => option.checked)
            .map(option => option.value);
        document.getElementById('hidden-company').value = selectedCompanies.join('|');
    }

    // ---------------- Primary Industry Filter ----------------
    const multiSelectIndustryContainer = document.querySelector('#primary-industry-filter');
    const multiSelectIndustrySelected = multiSelectIndustryContainer.querySelector('.multi-select-selected');
    const multiSelectIndustryDropdown = multiSelectIndustryContainer.querySelector('.multi-select-dropdown');
    const multiSelectIndustryOptions = multiSelectIndustryDropdown.querySelectorAll('.multi-select-option input[type="checkbox"]');
    const selectAllIndustryButton = multiSelectIndustryDropdown.querySelector('#select-all-industry');
    const clearAllIndustryButton = multiSelectIndustryDropdown.querySelector('#clear-all-industry');
    const applyIndustryButton = multiSelectIndustryDropdown.querySelector('#apply-selection-industry');
    const searchIndustryInput = multiSelectIndustryDropdown.querySelector('.multi-select-search input[type="text"]');

    // Toggle dropdown visibility for Primary Industry
    multiSelectIndustrySelected.addEventListener('click', function (event) {
        event.stopPropagation();  // Prevent event bubbling
        multiSelectIndustryDropdown.classList.toggle('show');
    });

    // Handle "Select All" for Primary Industry
    selectAllIndustryButton.addEventListener('click', function () {
        multiSelectIndustryOptions.forEach(option => option.checked = true);
        updateSelectedIndustryItems();
        filterForm.submit();
    });

    // Handle "Clear All" for Primary Industry
    clearAllIndustryButton.addEventListener('click', function () {
        multiSelectIndustryOptions.forEach(option => option.checked = false);
        updateSelectedIndustryItems();
        filterForm.submit();
    });

    // Apply button functionality for Primary Industry
    applyIndustryButton.addEventListener('click', function (event) {
        event.preventDefault();  // Prevent form submission
        copyIndustryFiltersToHiddenInputs();
        multiSelectIndustryDropdown.classList.remove('show');
        updateSelectedIndustryItems();
        filterForm.submit();
    });

    // Search functionality for Primary Industry
    searchIndustryInput.addEventListener('input', function () {
        let filter = this.value.toLowerCase();
        multiSelectIndustryOptions.forEach(option => {
            let label = option.nextElementSibling.textContent.toLowerCase();
            option.closest('.multi-select-option').style.display = label.includes(filter) ? '' : 'none';
        });
    });

    // Update selected items display for Primary Industry
    function updateSelectedIndustryItems() {
        const selectedItems = Array.from(multiSelectIndustryOptions)
            .filter(option => option.checked)
            .map(option => option.nextElementSibling.textContent);
        multiSelectIndustrySelected.innerHTML = selectedItems.join(', ') || 'Select Primary Industry';
    }

    // Copy selected industries to hidden input
    function copyIndustryFiltersToHiddenInputs() {
        const selectedIndustries = Array.from(multiSelectIndustryOptions)
            .filter(option => option.checked)
            .map(option => option.value);
        document.getElementById('hidden-industry-primary').value = selectedIndustries.join('|');
    }

    // Close dropdown when clicking outside
    document.addEventListener('click', function (e) {
        if (!multiSelectIndustryContainer.contains(e.target)) {
            multiSelectIndustryDropdown.classList.remove('show');
        }
    });

    // Restore selected values for Primary Industry on page load
    const hiddenIndustryInput = document.getElementById('hidden-industry-primary');
    if (hiddenIndustryInput && hiddenIndustryInput.value) {
        const selectedIndustryValues = hiddenIndustryInput.value.split('|');
        selectedIndustryValues.forEach(value => {
            const checkbox = document.querySelector(`#primary-industry-filter .multi-select-option input[type="checkbox"][value="${value}"]`);
            if (checkbox) checkbox.checked = true;
        });
        updateSelectedIndustryItems();
    }

    // -------------------- Relationship Filter Logic --------------------
    const relationshipContainer = document.querySelector('#relationship-filter');
    const relationshipDropdown = relationshipContainer.querySelector('.multi-select-dropdown');
    const relationshipSelected = relationshipContainer.querySelector('#relationship');
    const relationshipOptions = relationshipDropdown.querySelectorAll('.multi-select-option input[type="checkbox"]');
    const selectAllRelationshipButton = relationshipDropdown.querySelector('#select-all-relationship');
    const clearAllRelationshipButton = relationshipDropdown.querySelector('#clear-all-relationship');
    const applyRelationshipButton = relationshipDropdown.querySelector('.multi-select-apply');
    const searchRelationshipInput = relationshipDropdown.querySelector('.multi-select-search input[type="text"]');

    // Toggle dropdown visibility when the selected area is clicked (Relationship)
    relationshipSelected.addEventListener('click', function (event) {
        event.stopPropagation();  // Prevent the event from bubbling up
        relationshipDropdown.classList.toggle('show');
        console.log('Toggling Relationship dropdown');
    });


    // Apply button functionality for Relationship
    applyRelationshipButton.addEventListener('click', function (event) {
        event.preventDefault();
        copyRelationshipFiltersToHiddenInputs();
        relationshipDropdown.classList.remove('show');
        updateSelectedRelationshipItems();
        filterForm.submit();
    });


    // Update selected items display for Relationship
    function updateSelectedRelationshipItems() {
        const selectedItems = Array.from(relationshipOptions)
            .filter(option => option.checked)
            .map(option => option.nextElementSibling.textContent);
        relationshipSelected.innerHTML = selectedItems.join(', ') || 'Select Relationship';
    }

    // Copy selected relationships to hidden input
    function copyRelationshipFiltersToHiddenInputs() {
        const selectedRelationships = Array.from(relationshipOptions)
            .filter(option => option.checked)
            .map(option => option.value);
        document.getElementById('hidden-relationship').value = selectedRelationships.join('|');
    }

    // ---------------- Reset Button Logic ----------------
    const resetButton = document.getElementById('reset-filters-button');
    resetButton.addEventListener('click', function () {
        // Reset all checkboxes in the Country filter
        countryOptions.forEach(option => option.checked = false);
        updateSelectedCountryItems();

        // Reset all checkboxes in the Company filter
        companyOptions.forEach(option => option.checked = false);
        updateSelectedCompanyItems();

        // Reset all checkboxes in the Primary Industry filter
        multiSelectIndustryOptions.forEach(option => option.checked = false);
        updateSelectedIndustryItems();

        // Reset dropdowns or select fields (if any)
        document.getElementById('industry').value = '';
        document.getElementById('relationship').value = '';
        document.getElementById('revenue').value = '';
        document.getElementById('industry_primary').value = '';
        document.getElementById('industry_primary_sector').value = '';
        document.getElementById('combined_opp_score').value = '';
        document.getElementById('priority_status').value = '';
        document.getElementById('geography').value = '';

        // Clear hidden inputs
        document.getElementById('hidden-country').value = '';
        document.getElementById('hidden-company').value = '';
        document.getElementById('hidden-industry-primary').value = '';
        document.getElementById('hidden-relationship').value = '';

        // Optionally, submit the form after resetting
        filterForm.submit();  // Uncomment if you want the form to be submitted after reset
    });

    // -------------------- Close Dropdowns when clicking outside --------------------
    document.addEventListener('click', function (e) {
        // Close country dropdown
        if (!countryContainer.contains(e.target)) {
            countryDropdown.classList.remove('show');
        }

        // Close company dropdown
        if (!companyContainer.contains(e.target)) {
            companyDropdown.classList.remove('show');
        }

        // Close relationship dropdown
        if (!relationshipContainer.contains(e.target)) {
            relationshipDropdown.classList.remove('show');
        }
    });

    // -------------------- Restore Selected Values on Page Load --------------------
    const hiddenCountryInput = document.getElementById('hidden-country');
    if (hiddenCountryInput && hiddenCountryInput.value) {
        const selectedCountryValues = hiddenCountryInput.value.split('|');
        selectedCountryValues.forEach(value => {
            const checkbox = document.querySelector(`#country_${value}`);
            if (checkbox) {
                checkbox.checked = true;
            }
        });
        updateSelectedCountryItems();
    }

    const hiddenCompanyInput = document.getElementById('hidden-company');
    if (hiddenCompanyInput && hiddenCompanyInput.value) {
        // Split the stored company values (pipe-separated) into an array
        const selectedCompanyValues = hiddenCompanyInput.value.split(',');
    
        // Restore the checkboxes based on the selected values
        selectedCompanyValues.forEach(value => {
            const checkbox = document.querySelector(`.multi-select-container .multi-select-option input[type="checkbox"][value="${value}"]`);
            if (checkbox) {
                checkbox.checked = true;  // Check the checkbox
            }
        });
    
        // Display the restored selected company values
        document.getElementById('selected-company-display').innerHTML = selectedCompanyValues.length > 0 ? selectedCompanyValues.join(', ') : 'No companies selected';
    
        // Update the selected company items display
        updateSelectedCompanyItems();
    }

    const hiddenRelationshipInput = document.getElementById('hidden-relationship');
    if (hiddenRelationshipInput && hiddenRelationshipInput.value) {
        const selectedRelationshipValues = hiddenRelationshipInput.value.split('|');
        selectedRelationshipValues.forEach(value => {
            const checkbox = document.querySelector(`#relationship_${value}`);
            if (checkbox) {
                checkbox.checked = true;
            }
        });
        updateSelectedRelationshipItems();
    }
});

// Function to copy filter values from visible inputs to hidden inputs
function copyFiltersToHiddenInputs() {
    document.getElementById('hidden-industry').value = document.getElementById('industry').value;
    document.getElementById('hidden-revenue').value = document.getElementById('revenue').value;
    document.getElementById('hidden-combined-opp-score').value = document.getElementById('combined_opp_score').value;
    document.getElementById('hidden-priority-status').value = document.getElementById('priority_status').value;
    document.getElementById('hidden-geography').value = document.getElementById('geography').value;

    // Collect all selected countries
    const selectedCountries = Array.from(document.querySelectorAll('.multi-select-container #country .multi-select-option input[type="checkbox"]:checked'))
        .map(checkbox => checkbox.value);
    document.getElementById('hidden-country').value = selectedCountries.join('|');

    // Collect all selected companies
    const selectedCompanies = Array.from(document.querySelectorAll('.multi-select-container #company .multi-select-option input[type="checkbox"]:checked'))
        .map(checkbox => checkbox.value);
    document.getElementById('hidden-company').value = selectedCompanies.join('|');

    // Collect all selected relationships
    const selectedRelationships = Array.from(document.querySelectorAll('.multi-select-container #relationship .multi-select-option input[type="checkbox"]:checked'))
        .map(checkbox => checkbox.value);
    document.getElementById('hidden-relationship').value = selectedRelationships.join('|');

    // Collect all selected primary industry
    const selectedIndustryPrimary = Array.from(document.querySelectorAll('.multi-select-container #relationship .multi-select-option input[type="checkbox"]:checked'))
        .map(checkbox => checkbox.value);
    document.getElementById('hidden-industry-primary').value = selectedIndustryPrimary.join('|');
}




// Function to draw a bubble chart using Chart.js
function drawChart(chartData) {
    console.log("Chart Data: ", chartData);

    const ctx = document.getElementById('chart').getContext('2d');
    let minRevenue = Math.min(...chartData.map(d => parseFloat(d['Revenue (M)'])));
    let maxRevenue = Math.max(...chartData.map(d => parseFloat(d['Revenue (M)'])));
    
    // Prepare the chart data by calculating bubble sizes and mapping relevant fields
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
    const maxX = 20;
    const minY = 0;
    const maxY = 10;
    console.log('Axis values', minX, maxX, minY, maxY);

    // Chart data and configuration
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
                        color: '#000000' // Set x-axis tick color
                    },
                    min: minX,
                    max: maxX,
                    grid: {
                        display: false // Hide x-axis grid lines
                    }
                },
                y: {
                    ticks: {
                        color: '#000000' // Set y-axis tick color
                    },
                    min: minY,
                    max: maxY,
                    grid: {
                        display: false // Hide y-axis grid lines
                    }
                }
            },
            plugins: {
                tooltip: {
                    callbacks: {
                        // Custom tooltip label showing company and revenue
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
                    display: true,
                    align: 'center',
                    anchor: 'center',
                    color: 'white',
                    font: {
                        weight: 'bold',
                        size: 16
                    },
                    formatter: function(value, context) {
                        return value.company; // Show company name inside the bubble
                    }
                }
            }
        }
    };

    const myChart = new Chart(ctx, config); // Create and configure the chart
    myChart.update(); // Update the chart with the latest data
}
