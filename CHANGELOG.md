# Changelog

## [2.0.0](https://github.com/checkfirst-ltd/nestjs-outlook/compare/v1.0.0...v2.0.0) (2025-05-05)


### ⚠ BREAKING CHANGES

* **types:** Changed import source for Microsoft Graph types from '@microsoft/microsoft-graph-types' to local types. While functionally identical (re-exports), this change breaks type compatibility for library consumers who directly use these types.

### Features

* **types:** replace Microsoft Graph types with local re-exports ([2110d39](https://github.com/checkfirst-ltd/nestjs-outlook/commit/2110d39d601820bbece827aab262ee157e210f5a))

## 1.0.0 (2025-05-04)


### Features

* initial working module ([64ac682](https://github.com/checkfirst-ltd/nestjs-outlook/commit/64ac6820aa3ba8143bd9919db1d837992e999ec9))
