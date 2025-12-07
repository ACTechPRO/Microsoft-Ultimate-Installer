# 🚀 Microsoft 365 Ultimate Installer

> **Automated Installation & Licensing of Microsoft 365 Enterprise with Precision**

[![License](https://img.shields.io/badge/License-AC%20Tech%20Pro-blue)](LICENSE)
[![Version](https://img.shields.io/badge/Version-3.0.0-brightgreen)](https://github.com/ac-tech-pro/Microsoft-365-Ultimate-Installer/releases)
[![PowerShell](https://img.shields.io/badge/PowerShell-5.1%2B-0078D4)](https://docs.microsoft.com/en-us/powershell/)
[![Platform](https://img.shields.io/badge/Platform-Windows%2010%20%7C%2011-0078D4)](https://www.microsoft.com/windows)

---

## 📋 Visão Geral

O **Microsoft 365 Ultimate Installer** é uma solução automatizada robusta que simplifica a implementação do Microsoft 365 Enterprise em ambientes corporativos. Desenvolvido com PowerShell moderno, oferece uma experiência intuitiva através de uma interface WPF elegante, permitindo controle granular sobre cada aspecto da instalação.

### ✨ Principais Características

| Recurso | Descrição |
|---------|-----------|
| 🎯 **Modo Express** | Instalação rápida do pacote padrão em um clique |
| 🎨 **Interface WPF** | Dashboard interativo com preview em tempo real |
| 🌍 **Multi-idioma** | Suporte a 29+ idiomas (EN, PT-BR, ES, JA, DE, FR, ZH, IT, KO, RU, etc.) |
| 🔒 **Privacidade Máxima** | Telemetria desabilitada, sem rastreamento |
| ⚙️ **Customização Total** | Controle sobre cada aplicação instalada |
| 🛡️ **Ativação Automática** | HWID (Windows) + Ohook (Office) - sem intervenção |
| 📦 **Aplicações Incluídas** | Word, Excel, PowerPoint, Outlook, Teams, Clipchamp, Project Pro, Visio Pro |
| 🧹 **Limpeza Automática** | Remove instalações antigas e arquivos temporários |
| 📊 **Logging Detalhado** | Rastreamento completo salvo na Desktop |
| 🔄 **Resiliência** | Tratamento robusto de erros com rollback automático |

---

## 🎯 Aplicações Disponíveis

### 📌 Pacote Express (Padrão)
- ✅ Microsoft Word
- ✅ Microsoft Excel
- ✅ Microsoft PowerPoint
- ✅ Microsoft Outlook
- ✅ Microsoft Teams
- ✅ Clipchamp (Editor de vídeo)

### 📌 Pacote Profissional (Adicional)
- 📊 Access
- 📋 Publisher
- 🗂️ Project Pro
- 🏗️ Visio Pro
- ⚡ Power Automate Desktop

### 🚫 Aplicações Excluídas
- OneDrive
- OneNote
- Lync
- Groove (Música)

---

## 🛠️ Requisitos Técnicos

### Mínimos
- **OS**: Windows 10 (Build 1909+) ou Windows 11
- **Arquitetura**: 64-bit
- **PowerShell**: 5.1 ou superior
- **Memória RAM**: 2 GB mínimo (4 GB recomendado)
- **Espaço em Disco**: 8 GB disponível
- **Conexão de Internet**: Obrigatória (>2 Mbps recomendado)
- **Privilégios**: Administrador

### Recomendados
- **OS**: Windows 11 (versão recente)
- **Processador**: Multi-core (2.5 GHz+)
- **Memória RAM**: 8 GB+
- **SSD**: Para desempenho otimizado

---

## 📥 Instalação & Uso

### 1️⃣ Baixar o Script

```powershell
# Clone o repositório ou baixe diretamente
git clone https://github.com/ac-tech-pro/Microsoft-365-Ultimate-Installer.git
```

### 2️⃣ Abrir como Administrador

```powershell
# Navegue até a pasta do script
cd .\Microsoft-365-Ultimate-Installer

# Execute com privilégios de administrador
.\Microsoft 365 Ultimate Installer.ps1
```

### 3️⃣ Escolher Modo de Instalação

#### 🚀 Modo Express
- Clique em **"Express Installation"**
- Aplicações padrão serão instaladas automaticamente
- Tempo estimado: 15-30 minutos

#### ⚙️ Modo Customizado
- Clique em **"Custom Installation"**
- Selecione cada aplicação desejada
- Configure idiomas (até 4 idiomas simultâneos)
- Configure canal de atualização
- Inicie a instalação

### 4️⃣ Acompanhar o Progresso

- ✅ Barra de progresso em tempo real
- 📊 Detalhes de cada etapa
- 🔔 Notificações de status
- ⏱️ Tempo estimado restante

---

## 🔧 Arquitetura Técnica

### Fluxo de Execução

```
┌─────────────────────────────────────────┐
│  Inicialização & Validação de Ambiente  │
├─────────────────────────────────────────┤
│  Detecção de Idioma & Configuração UI   │
├─────────────────────────────────────────┤
│  Interface WPF (Janela Principal)       │
├─────────────────────────────────────────┤
│  Coleta de Preferências do Usuário      │
├─────────────────────────────────────────┤
│  Download ODT & Configuração            │
├─────────────────────────────────────────┤
│  Limpeza de Instalações Antigas         │
├─────────────────────────────────────────┤
│  Instalação via Office Deployment Tool  │
├─────────────────────────────────────────┤
│  Instalação de Complementos (Winget)    │
├─────────────────────────────────────────┤
│  Ativação Automática (HWID + Ohook)     │
├─────────────────────────────────────────┤
│  Limpeza de Arquivos Temporários        │
├─────────────────────────────────────────┤
│  Relatório Final & Logging              │
└─────────────────────────────────────────┘
```

### Componentes Principais

#### 🔹 Camada de UI (WPF XAML)
- Interface responsiva e moderna
- Tema adaptativo Windows 10/11
- Indicadores visuais de progresso
- Feedback em tempo real

#### 🔹 Engine de Instalação
- Microsoft Office Deployment Tool (ODT)
- Configuração XML dinâmica
- Suporte a múltiplos canais (Current, Deferred, Semi-Annual)
- Download inteligente com retry automático

#### 🔹 Sistema de Ativação
- **HWID**: Ativação automática do Windows
- **Ohook**: Ativação automática do Office
- Scripts MAS (Microsoft Activation Scripts)
- Validação pós-ativação

#### 🔹 Logging & Diagnostics
- Arquivo de log detalhado na Desktop
- Timestamps precisos para cada operação
- Stack traces para troubleshooting
- Relatório de erros estruturado

#### 🔹 Runspaces PowerShell
- Múltiplas threads de execução
- UI nunca congela durante instalação
- Processamento paralelo de tarefas
- Sincronização de estado thread-safe

---

## 📝 Parâmetros de Execução

```powershell
# Execução padrão
.\Microsoft 365 Ultimate Installer.ps1

# Forçar limpeza de mutex (se preso)
.\Microsoft 365 Ultimate Installer.ps1 -Force

# Execução direta (sem relaunch oculto)
.\Microsoft 365 Ultimate Installer.ps1 -IsHidden
```

---

## 🌐 Idiomas Suportados

O script automaticamente detecta o idioma do Windows e oferece suporte completo a:

| Categoria | Idiomas |
|-----------|---------|
| **Europeus** | Inglês (EN-US, EN-GB), Português (PT-BR), Espanhol (ES-ES), Francês (FR-FR), Alemão (DE-DE), Italiano (IT-IT) |
| **Asiáticos** | Japonês (JA-JP), Chinês Simplificado (ZH-CN), Coreano (KO-KR) |
| **Eslávicos** | Russo (RU-RU), Polonês (PL-PL), Checo (CS-CZ) |
| **Nórdicos** | Sueco (SV-SE), Dinamarquês (DA-DK), Norueguês (NB-NO) |
| **Outros** | Holandês (NL-NL), Turco (TR-TR), Grego (EL-GR), Hebraico (HE-IL), Árabe (AR-SA) |

---

## 🔐 Segurança & Privacidade

### Medidas de Segurança

✅ **Execução Oculta**
- Nenhuma janela visível durante execução
- Minimiza suspeita de antivírus
- Processo transparente ao usuário

✅ **Validação de Integridade**
- Verificação de hash de downloads
- Validação de assinatura de scripts
- Rollback em caso de falha

✅ **Privacidade Máxima**
- Telemetria desabilitada
- Rastreamento desativado
- Conexões de diagnóstico bloqueadas
- Cortana desabilitado

### Dados de Telemetria

🚫 **Desabilitados**
- Diagnostic Data Collection
- Customer Experience Improvement Program (CEIP)
- Connected Experiences
- Microsoft Consumer Experiences

---

## 📊 Estrutura de Arquivos

```
Microsoft-365-Ultimate-Installer/
├── Microsoft 365 Ultimate Installer.ps1 (Script Principal)
├── README.md (Este arquivo)
├── LICENSE (Propriedade AC Tech Pro)
└── [Arquivos temporários durante execução]
    ├── %LOCALAPPDATA%\Temp\M365Ultimate_Installation\
    └── Desktop\Microsoft 365 Ultimate Installer.log
```

---

## 🚨 Troubleshooting

### ❌ Problema: "Another instance is already running"

```powershell
# Solução 1: Use o parâmetro -Force
.\Microsoft 365 Ultimate Installer.ps1 -Force

# Solução 2: Limpe manualmente o mutex
Get-Process powershell -ErrorAction SilentlyContinue | Stop-Process -Force
```

### ❌ Problema: Falha de Download

```powershell
# Verificar conectividade
Test-NetConnection -ComputerName download.microsoft.com -Port 443

# Verificar firewall
Get-NetFirewallProfile | Select-Object Name, Enabled
```

### ❌ Problema: Ativação não funcionou

```powershell
# Verificar status de ativação
slmgr /xpr-status  # Windows
ospp.vbs /dstat    # Office
```

---

## 📜 Changelog

### v3.0.0 (Current)
- ✨ Script unificado em um único arquivo
- 🔒 Execução oculta nativa via PowerShell (sem P/Invoke)
- 🎯 Interface WPF melhorada
- 🌍 Suporte a 29+ idiomas
- 🛡️ Ativação automática HWID + Ohook
- 📊 Sistema de logging estruturado
- 🔄 Tratamento robusto de erros

---

## 📞 Suporte & Contacto

**AC Tech Pro**
- 🌐 [GitHub Organization](https://github.com/ac-tech-pro)
- 📧 Suporte via Issues do repositório
- 🔗 [Repositório Principal](https://github.com/ac-tech-pro/Microsoft-365-Ultimate-Installer)

---

## ⚖️ Licença

**Propriedade privada de AC Tech Pro** © 2025

Este projeto é mantido e licenciado exclusivamente para AC Tech Pro. Uso não autorizado é proibido.

---

## 🤝 Contribuindo

Contribuições são bem-vindas! Por favor:

1. Fork o repositório
2. Crie uma branch para sua feature (`git checkout -b feature/AmazingFeature`)
3. Commit suas mudanças (`git commit -m 'Add AmazingFeature'`)
4. Push para a branch (`git push origin feature/AmazingFeature`)
5. Abra um Pull Request

---

<div align="center">

**Desenvolvido com ❤️ por AC Tech Pro**

[⬆ Voltar ao topo](#-microsoft-365-ultimate-installer)

</div>
