window.PraxisBrand = {
  logo: "assets/praxis-logo.jpg",
  sedeCentro: "assets/sede-centro-real.jpg"
};

document.addEventListener("DOMContentLoaded", () => {
  document.querySelectorAll("[data-brand-logo]").forEach(img => img.src = window.PraxisBrand.logo);
  document.querySelectorAll("[data-brand-school]").forEach(img => img.src = window.PraxisBrand.sedeCentro);
});
