# TeamFoundationServer

[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)

Utility scripts and tools for Team Foundation Server (TFS) and Azure DevOps Server, providing migration and management capabilities.

## Overview

This repository contains PowerShell scripts that help with TFS management tasks, particularly focusing on migration to Git. These utilities are designed to be auditable, secure, and suitable for use in enterprise environments.

## Utilities

### ConvertTo-Git.ps1

A comprehensive TFVC to Git migration tool that accurately preserves your project history by replaying all changes chronologically.

**Key Features:**

- Recursively processes all changesets through the entire branch hierarchy
- Preserves complete commit history with original timestamps and authors
- Handles all TFVC change types (add, edit, delete, rename, branch)
- Supports secure authentication methods
- Optimized for pipeline execution in Azure DevOps
- Creates a flat migration for consistent handling of large projects

**When to use this approach:**

While this migration approach can be slower than other methods, it provides several advantages:
- Transparent, auditable conversion process with simple, reviewable code
- Consistent handling of complex projects with many stale branches
- Complete preservation of history, including changesets, timestamps, and authors
- Suitable for regulated environments requiring migration validation

**Post-migration:**

After migration to Git, project leads can:
- Split the monolithic repository into appropriate Git branches as needed
- Implement proper Git workflows optimized for their team's needs
- Take advantage of Git's distributed capabilities

## Usage

### Authentication Options

The script supports several authentication methods:

```powershell
# Windows Authentication
.\ConvertTo-Git.ps1 -TfsProject "$/ProjectName" -OutputPath "C:\OutputFolder" -TfsCollection "https://Some.Private.Server/tfs/DefaultCollection"

# Username/Password Authentication
.\ConvertTo-Git.ps1 -TfsProject "$/ProjectName" -OutputPath "C:\OutputFolder" -TfsCollection "https://Some.Private.Server/tfs/DefaultCollection" -TfsUserName "your_username" -TfsPassword "your_password"

# With password from environment variable
.\ConvertTo-Git.ps1 -TfsProject "$/ProjectName" -OutputPath "C:\OutputFolder" -TfsCollection "https://Some.Private.Server/tfs/DefaultCollection" -TfsUserName "your_username"
```

### Pipeline Integration

Easily integrate with Azure DevOps pipelines:

```yaml
steps:
- task: PowerShell@2
  inputs:
    filePath: '.\ConvertTo-Git.ps1'
    arguments: '-TfsProject "$/YourProject" -OutputPath "$(Build.ArtifactStagingDirectory)" -TfsCollection "https://dev.azure.com/yourorg" -TfsUserName "$(TfsUserName)" -TfsPassword "$(TfsPassword)"'
  displayName: 'Convert TFVC to Git'
```

## Requirements

- Windows PowerShell 5.1 or newer
- Visual Studio with Team Explorer (2019 or 2022) installed
- Git command-line tools
- Appropriate permissions in the TFS/Azure DevOps project

## Security Notes

- Credentials are handled securely using SecureString objects
- Passwords are cleared from memory after use
- Compatible with Azure DevOps pipeline secret variables
- No credentials are logged or displayed in plain text

## License

This project is licensed under the MIT License - see the LICENSE file for details.

## Contributing

Contributions are welcome! Please feel free to submit a Pull Request.