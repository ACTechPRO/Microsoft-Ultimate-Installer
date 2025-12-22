<div align="center">
  <img src="assets/icon.png" alt="Microsoft Ultimate Installer Icon" width="128" />
  <h1>Microsoft Ultimate Installer</h1>
  <p>
    <b>Automação, Controle e Elegância.</b><br>
    A solução definitiva para gerenciamento de softwares Microsoft e otimização do Windows.
  </p>
</div>

---

## 🚀 Sobre o Projeto

O **Microsoft Ultimate Installer** é uma ferramenta PowerShell avançada com interface gráfica moderna (WPF/XAML) projetada para facilitar a instalação, ativação e limpeza de produtos Microsoft. Focado em privacidade e eficiência, ele elimina a necessidade de múltiplos instaladores e configurações manuais.

## ✨ Características Principais

| Recurso | Detalhes |
| :--- | :--- |
| **🎨 Interface Premium** | Design moderno, tema escuro, janelas redimensionáveis e centralizadas. |
| **🧹 Deep Clean** | **Novo:** Desinstalação silenciosa e completa de VS (todas as versões), Office, Teams e Apps. Inclui limpeza agressiva de atalhos e residuais. |
| **🔇 Instalação Silenciosa** | **Novo:** Instalação e desinstalação do Visual Studio sem popups (`--quiet`), garantindo fluxo ininterrupto. |
| **🛡️ Ativação Inteligente** | Processos automáticos de licenciamento (HWID / Ohook) sem intervenção do usuário. |
| **⚡ Performance** | Instalação otimizada via Winget (com `--disable-interactivity`) e setups offline. Bloqueio de auto-início de apps. |
| **🔒 Privacidade Total** | Telemetria desativada por padrão. Sem rastreamento de uso. |

## 🛠️ Funcionalidades

### Instalação e Configuração
*   **Microsoft 365 / Office**: Instalação personalizada (Word, Excel, PowerPoint, Project, Visio).
*   **Visual Studio**: Instalação automática da versão Enterprise (Insiders) com cargas de trabalho selecionadas.
*   **Ferramentas Essenciais**: VS Code, PowerToys, Microsoft Teams, UniGetUI.
*   **Windows 10/11**: Scripts de otimização e debloat integrados.

### Manutenção e Remoção
*   **Complete Removal (Modo Uninstall)**:
    *   Detecta e remove todas as instâncias do Visual Studio via `vswhere`.
    *   Itera e remove apps instalados via Store (Appx) e Win32 (Winget).
    *   **Limpeza de Atalhos**: Varredura ativa no Desktop e Menu Iniciar para remover ícones "fantasmas" pós-desinstalação.
    *   Limpeza profunda de arquivos temporários e registros.

## 📋 Requisitos do Sistema

*   **SO**: Windows 10 (1809+) ou Windows 11.
*   **PowerShell**: Versão 5.1 ou superior.
*   **Permissões**: Privilégios de Administrador (obrigatório).
*   **Internet**: Conexão estável para download dos pacotes.

## 🚀 Como Usar

1.  Baixe o repositório ou o arquivo `.ps1`.
2.  Abra o **PowerShell** como Administrador.
3.  Execute o script:

```powershell
.\Microsoft Ultimate Installer.ps1
```

> **Nota:** Se houver restrições de execução de script, use: `Set-ExecutionPolicy Unrestricted -Scope Process` antes de executar.

## 🔗 Links Úteis

*   [Repositório GitHub](https://github.com/ACTechPRO/Microsoft-Ultimate-Installer)
*   [Relatar Problemas (Issues)](https://github.com/ACTechPRO/Microsoft-Ultimate-Installer/issues)

---

<div align="center">
  <sub>Desenvolvido por <b>AC Tech</b> • Brasil 🇧🇷</sub>
</div>
