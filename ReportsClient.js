/**
 * ==========================================
 * Initialize Reports Module
 * ==========================================
 */
function initializeReports() {

  console.log("Reports Module Loaded");

  const today = new Date().toISOString().split("T")[0];

  const fromDate = document.getElementById("reportFromDate");

  const toDate = document.getElementById("reportToDate");

  // Default Dates
  fromDate.value = today;
  toDate.value = today;

  // Restrict Future Dates
  fromDate.max = today;
  toDate.max = today;

  // To Date cannot be less than From Date
  toDate.min = today;

  // From Date Changed
  fromDate.addEventListener("change", function () {

    toDate.min = this.value;

    if (toDate.value < this.value) {

      toDate.value = this.value;

    }

  });

  // To Date Changed
  toDate.addEventListener("change", function () {

    if (this.value > today) {

      this.value = today;

    }

  });

  // Load Today's KPI
  google.script.run
    .withSuccessHandler(function (data) {

      console.log("Server Response:", data);

      renderReportDashboard(data);

    })

    .withFailureHandler(function (error) {

      console.error(error);

    })

    .getReportsData(today, today);
}


/**
 * ==========================================
 * Generate Report
 * ==========================================
 */
function generateReport() {

  const fromDate =
    document.getElementById("reportFromDate").value;

  const toDate =
    document.getElementById("reportToDate").value;

  if (!fromDate || !toDate) {

    alert("Please select both dates.");

    return;

  }

  if (fromDate > toDate) {

    alert("From Date cannot be greater than To Date.");

    return;

  }

  google.script.run

    .withSuccessHandler(function (response) {

      console.log("Reports Response:", response);

      console.log("Orders:", response.orders);

      console.log("Orders Count:",
        response.orders ? response.orders.length : 0);

      if (!response.success) {

        alert(response.message);

        return;

      }

      renderSummary(response);

      renderProfitSummary(response);

      renderTopSellingItems(response.topSellingItems);

      renderOrderTable(response.orders);

    })

    .withFailureHandler(function (error) {

      console.error(error);

      alert(error);

    })

    .getReportsData(fromDate, toDate);

}


/**
 * ==========================================
 * Render Summary
 * ==========================================
 */
function renderSummary(data) {

  console.log("===== NEW renderSummary =====");

  console.log(data);

  debugger;

}




function renderOrderTable(orders) {

  const tbody = document.getElementById("reportOrdersBody");

  tbody.innerHTML = "";

  if (!orders || orders.length === 0) {

    tbody.innerHTML = `
            <tr>
                <td colspan="7" class="text-center text-muted">
                    No records found.
                </td>
            </tr>
        `;

    return;

  }

  orders.forEach(function (order) {

    tbody.innerHTML += `

            <tr>

                <td>${order.orderId}</td>

                <td>${order.orderDate}</td>

                <td>${order.customer}</td>

                <td>${order.mobile}</td>

                <td>${order.paymentMode}</td>

                <td>${order.paymentStatus}</td>

                <td class="text-end">₹${Number(order.grandTotal).toFixed(2)}</td>

            </tr>

        `;

  });

}