# Contributing to NestJS Outlook

Thank you for considering contributing to the NestJS Outlook module! This document provides guidelines and instructions for contributing.

## Code of Conduct

Please be respectful and constructive in all interactions related to this project.

## Brand Guidelines

When referencing this library or Checkfirst in your contributions:

- Always refer to the library as "@checkfirst/nestjs-outlook" or "NestJS Outlook"
- The company name should always be written as "Checkfirst" (not "CheckFirst" or "Check First")
- Use the official logo when creating visual materials related to the library
- Main brand colors:
  - Primary Blue: #0f4266
  - Secondary Teal: #42b6b3
- For more detailed brand guidelines, please contact [info@checkfirst.ai](mailto:info@checkfirst.ai)

## How to Contribute

1. Fork the repository
2. Create a new branch (`git checkout -b feature/your-feature-name`)
3. Make your changes
4. Run tests (`npm test`)
5. Commit your changes (`git commit -m 'Add some feature'`)
6. Push to your branch (`git push origin feature/your-feature-name`)
7. Create a new Pull Request

## Development Setup

1. Clone the repository
2. Install dependencies: `npm install`
3. Build the project: `npm run build`
4. Run tests: `npm test`

## Pull Request Guidelines

- Use conventional commit format for PR titles (e.g., `feat: add new authentication method`, `fix: resolve login issue`)
- Include a reference to related GitHub issues if applicable (e.g., `feat: add new authentication method (#123)`)
- Include a clear description of the changes
- Update documentation if necessary
- Add tests for new features
- Follow the existing code style
- Make sure all tests pass

### Conventional Commit Format

PR titles should follow the [Conventional Commits](https://www.conventionalcommits.org/) specification:

```
<type>[optional scope]: <description>
```

Common types include:
- feat: A new feature
- fix: A bug fix
- docs: Documentation changes
- style: Changes that don't affect code functionality (formatting, etc.)
- refactor: Code changes that neither fix bugs nor add features
- test: Adding or fixing tests
- chore: Changes to build process or auxiliary tools

When referencing issues, append the issue number in parentheses to the title:

```
feat: implement OAuth authentication (#42)
```

## Testing

Run tests with:

```bash
npm test
```

## Linting

Run linting with:

```bash
npm run lint
```

## Building

Build the project with:

```bash
npm run build
```

## Questions?

If you have any questions, please open an issue on GitHub. 