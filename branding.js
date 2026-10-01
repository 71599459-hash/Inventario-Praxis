window.PraxisBrand = {
  logo: "https://app-auth.app.praxis.edu.pe/asset/image/logo/colegio-praxis.png",
  logoFallback: "praxis-logo-transparent.png?v=20261001-1",
  sedeCentro: "sede-centro.avif?v=20261001-1"
};

document.addEventListener("DOMContentLoaded", () => {
  document.querySelectorAll("[data-brand-logo]").forEach((img) => {
    img.decoding = "async";
    img.src = window.PraxisBrand.logo;
    img.onerror = () => {
      img.onerror = null;
      img.src = window.PraxisBrand.logoFallback;
    };
  });

  document.querySelectorAll("[data-brand-school]").forEach((img) => {
    img.decoding = "async";
    img.src = window.PraxisBrand.sedeCentro;
  });
});
