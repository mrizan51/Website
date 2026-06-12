/* Premier Exports International — site behaviour */
(function () {
  "use strict";

  /* ----- Mobile navigation ----- */
  var toggle = document.querySelector(".nav-toggle");
  if (toggle) {
    toggle.addEventListener("click", function () {
      var open = document.body.classList.toggle("nav-open");
      toggle.setAttribute("aria-expanded", open ? "true" : "false");
    });
    document.addEventListener("click", function (e) {
      if (
        document.body.classList.contains("nav-open") &&
        !e.target.closest(".site-header")
      ) {
        document.body.classList.remove("nav-open");
        toggle.setAttribute("aria-expanded", "false");
      }
    });
  }

  /* ----- Reveal on scroll ----- */
  var reveals = document.querySelectorAll(".reveal");
  if ("IntersectionObserver" in window && reveals.length) {
    var io = new IntersectionObserver(
      function (entries) {
        entries.forEach(function (entry) {
          if (entry.isIntersecting) {
            entry.target.classList.add("in");
            io.unobserve(entry.target);
          }
        });
      },
      { threshold: 0.12, rootMargin: "0px 0px -40px 0px" }
    );
    reveals.forEach(function (el) { io.observe(el); });
  } else {
    reveals.forEach(function (el) { el.classList.add("in"); });
  }

  /* ----- Quote form → pre-filled email ----- */
  var form = document.getElementById("quote-form");
  if (form) {
    form.addEventListener("submit", function (e) {
      e.preventDefault();
      if (!form.checkValidity()) {
        form.reportValidity();
        return;
      }
      var v = function (id) {
        var el = document.getElementById(id);
        return el && el.value ? el.value.trim() : "";
      };
      var subject = "Quote request — " + v("qf-products") + " — " + v("qf-company");
      var body = [
        "Name: " + v("qf-name"),
        "Company: " + v("qf-company"),
        "Email: " + v("qf-email"),
        "Country / Market: " + v("qf-country"),
        "Products of interest: " + v("qf-products"),
        "",
        v("qf-message"),
        "",
        "— Sent from premierexport.in"
      ].join("\n");
      var mailto =
        "mailto:" + (window.PEI_CONTACT_EMAIL || "premier.pei@gmail.com") +
        "?subject=" + encodeURIComponent(subject) +
        "&body=" + encodeURIComponent(body);
      window.location.href = mailto;
      var status = document.getElementById("form-status");
      if (status) {
        status.textContent = "Your email app should now open with the enquiry pre-filled.";
        status.classList.add("show");
      }
    });
  }
})();
