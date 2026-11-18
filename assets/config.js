/* =========================================================================
   DSL Lab Submission Portal: site configuration
   After deploying backend/Code.gs as a Google Apps Script Web App,
   paste its URL below (it ends with /exec), commit, and you're live.
   While API_URL is empty, the site runs in PREVIEW mode (nothing is saved).
   ========================================================================= */
window.DSL_CONFIG = {
  API_URL: 'https://script.google.com/macros/s/AKfycbyjfqQDfWNrdR4wP7nyqfTiRCAn4dHiGl4wByF-ngyBhvPy7OiPc_uicbXsPsC_j_vo/exec',

  // Divisions shown in the portal (TE research papers: A to E only)
  DIVISIONS: ['A', 'B', 'C', 'D', 'E'],

  // University logo: a local copy (assets/logo.png) is used first if you add one,
  // otherwise the official logo from adypu.edu.in is shown.
  LOGO_LOCAL: 'assets/logo.png',
  LOGO_REMOTE: 'https://adypu.edu.in/wp-content/themes/adypu/images/logo.png'
};
