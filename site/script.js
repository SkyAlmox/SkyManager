const menuToggle = document.getElementById('menuToggle');
const mainNav = document.getElementById('mainNav');
const overlay = document.getElementById('overlay');
const submenuBtn = document.getElementById('submenuBtn');
const subMenu = document.getElementById('subMenu');

function toggleMenu() {
    menuToggle.classList.toggle('is-active');
    mainNav.classList.toggle('is-active');
    overlay.classList.toggle('is-active');
    document.body.classList.toggle('no-scroll');
}

function toggleSubmenu(e) {
    e.preventDefault();
    subMenu.classList.toggle('is-open');
    submenuBtn.classList.toggle('is-open');
}

menuToggle.addEventListener('click', toggleMenu);
overlay.addEventListener('click', toggleMenu);
submenuBtn.addEventListener('click', toggleSubmenu);

const navLinks = document.querySelectorAll('.menu-item a:not(.has-submenu > .menu-link-wrapper > a)');
navLinks.forEach(link => {
    link.addEventListener('click', () => {
        if (mainNav.classList.contains('is-active')) toggleMenu();
    });
});