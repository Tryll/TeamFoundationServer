# TeamFoundationServer

[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)

Utility scripts and tools for Team Foundation Server (TFS) and Azure DevOps Server, providing migration and management capabilities.

## Overview

This repository contains PowerShell scripts that help with TFS management tasks, particularly focusing on migration to Git. These utilities are designed to be auditable, secure, and suitable for use in enterprise environments.

## Utilities

### ConvertTo-GitFlat.ps1

A comprehensive TFVC to Git migration tool that accurately preserves your project history by replaying all changes chronologically, in less than 400 lines of code.

**Key Features:**

- Creates a unified source tree for the complete TFS project, for consistent handling of large projects
- TFS branches are processed as directories (owners will split to the required branches after migration)
- Recursively processes all changesets through the entire project hierarchy
- Preserves complete commit history with original timestamps and authors
- Handles all TFVC change types (add, edit, delete, rename, branch)
- Supports secure authentication methods
- Optimized for pipeline execution in Azure DevOps

**When to use this approach:**

While this migration approach can be slower than other methods, it provides several advantages:
- Transparent, auditable conversion process with simple, reviewable code
- Consistent handling of complex projects with many stale and deleted branches
- Complete preservation of history, including changesets, timestamps, and authors
- Suitable for regulated environments requiring migration validation

**Post-migration:**

After migration to Git, project leads can:
- Split the monolithic repository into appropriate Git branches as needed
- Implement proper Git workflows optimized for their team's needs
- Take advantage of Git's distributed capabilities


**Real world run:**
```code
2025-03-16T05:37:53.7499263Z Conversion completed!
2025-03-16T05:37:53.7500659Z Total changesets processed: 4928
2025-03-16T05:37:53.7500916Z Total files processed: 271493
2025-03-16T05:37:53.7501140Z Total conversion time: 6 hours, 26 minutes, 18 seconds
```

## License

This project is licensed under the MIT License - see the LICENSE file for details.

## Contributing

Contributions are welcome! Please feel free to submit a Pull Request.