function NavBar() {
  return (
    <nav className="navbar navbar-expand-lg bg-body-tertiary fixed-top">
      <div className="container-fluid d-flex justify-content-between align-items-center">
        <a className="navbar-brand" href="#">
          <img
            src="/logo-branca.png"
            alt="Logo"
            className="img-fluid d-inline-block align-text-top"
            style={{ maxHeight: '40px' }}
          />
        </a>
        <div className="collapse navbar-collapse justify-content-center" id="navbarNavDropdown">
          <ul className="navbar-nav mx-auto">
            <li className="nav-item">
              <a className="nav-link" href="#">Identidade Visual</a>
            </li>
            <li className="nav-item dropdown">
              <a
                className="nav-link dropdown-toggle"
                href="#"
                role="button"
                data-bs-toggle="dropdown"
                aria-expanded="false"
              >
                Produtos
              </a>
              <ul className="dropdown-menu">
                <li><a className="dropdown-item" href="#">Carregadores de Celular</a></li>
                <li><a className="dropdown-item" href="#">Comunicação Visual</a></li>
                <li><a className="dropdown-item" href="#">Estruturas para Mídia OOH</a></li>
                <li><a className="dropdown-item" href="#">Marcenaria</a></li>
                <li><a className="dropdown-item" href="#">Projetos Especiais de Arquitetura</a></li>
                <li><a className="dropdown-item" href="#">Totens Digitais e Interativos</a></li>
              </ul>
            </li>
            <li className="nav-item">
              <a className="nav-link" href="#">Vamos Conversar ?</a>
            </li>
          </ul>
        </div>
        <div className="d-flex">
          <a href="https://instagram.com/seuPerfil" target="_blank" rel="noopener noreferrer" className="nav-link">
            <i className="bi bi-instagram" style={{ fontSize: '1.5rem' }}></i>
          </a>
          <a href="https://wa.me/seuNumero" target="_blank" rel="noopener noreferrer" className="nav-link ms-3">
            <i className="bi bi-whatsapp" style={{ fontSize: '1.5rem' }}></i>
          </a>
        </div>
      </div>
    </nav>
  );
}

export default NavBar;