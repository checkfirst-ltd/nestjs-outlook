# Changelog

## [5.2.0](https://github.com/checkfirst-ltd/nestjs-outlook/compare/v5.1.3...v5.2.0) (2026-01-28)


### Features

* add update event to act as upsert based on trasnaction id ([#71](https://github.com/checkfirst-ltd/nestjs-outlook/issues/71)) ([0d5abf0](https://github.com/checkfirst-ltd/nestjs-outlook/commit/0d5abf071e91d585e156b6e2a40b62a480276fc5))

## [5.1.3](https://github.com/checkfirst-ltd/nestjs-outlook/compare/v5.1.2...v5.1.3) (2026-01-19)


### Bug Fixes

* Delta Link Expired with error stateNotFound ([#69](https://github.com/checkfirst-ltd/nestjs-outlook/issues/69)) ([0f24e87](https://github.com/checkfirst-ltd/nestjs-outlook/commit/0f24e87901240c34377617ab5ebf8a6746552900))

## [5.1.2](https://github.com/checkfirst-ltd/nestjs-outlook/compare/v5.1.1...v5.1.2) (2026-01-18)


### Bug Fixes

* delete event with exponential backoff and observability ([#67](https://github.com/checkfirst-ltd/nestjs-outlook/issues/67)) ([03e35d6](https://github.com/checkfirst-ltd/nestjs-outlook/commit/03e35d667d53f4ca073d16759ce8e2a360415ad7))

## [5.1.1](https://github.com/checkfirst-ltd/nestjs-outlook/compare/v5.1.0...v5.1.1) (2026-01-14)


### Bug Fixes

* correct user ID handling in webhook and delta sync operations ([#63](https://github.com/checkfirst-ltd/nestjs-outlook/issues/63)) ([442a884](https://github.com/checkfirst-ltd/nestjs-outlook/commit/442a8840101d17908fa4d1a19460d5ac191ac07f))

## [5.1.0](https://github.com/checkfirst-ltd/nestjs-outlook/compare/v5.0.1...v5.1.0) (2026-01-13)


### Features

* **Calendar:** expose the event status types for host app usage ([#64](https://github.com/checkfirst-ltd/nestjs-outlook/issues/64)) ([2ceacae](https://github.com/checkfirst-ltd/nestjs-outlook/commit/2ceacae7496ef14f4f04987e2ded3c8e0b6e21a7))

## [5.0.1](https://github.com/checkfirst-ltd/nestjs-outlook/compare/v5.0.0...v5.0.1) (2026-01-12)


### Bug Fixes

* Fix/early checkings ([#61](https://github.com/checkfirst-ltd/nestjs-outlook/issues/61)) ([7dccd6a](https://github.com/checkfirst-ltd/nestjs-outlook/commit/7dccd6a4013da6be6078a183a9fd8f86175f523f))

## [5.0.0](https://github.com/checkfirst-ltd/nestjs-outlook/compare/v4.7.0...v5.0.0) (2026-01-11)


### ⚠ BREAKING CHANGES

* get active subscription by external user id ([#58](https://github.com/checkfirst-ltd/nestjs-outlook/issues/58))

### Features

* get active subscription by external user id ([#58](https://github.com/checkfirst-ltd/nestjs-outlook/issues/58)) ([29a7463](https://github.com/checkfirst-ltd/nestjs-outlook/commit/29a7463a57a33f69df0e1b7bf536c2b27dcefc81))

## [4.7.0](https://github.com/checkfirst-ltd/nestjs-outlook/compare/v4.6.1...v4.7.0) (2026-01-11)


### Features

* handling lifecycle events ([#56](https://github.com/checkfirst-ltd/nestjs-outlook/issues/56)) ([c3c49fb](https://github.com/checkfirst-ltd/nestjs-outlook/commit/c3c49fb4d7ee40b0b1fa3a8c6a0f8ac01d91f937))

## [4.6.1](https://github.com/checkfirst-ltd/nestjs-outlook/compare/v4.6.0...v4.6.1) (2025-12-24)


### Bug Fixes

* minor bug in odata.type comparison ([#54](https://github.com/checkfirst-ltd/nestjs-outlook/issues/54)) ([617eba5](https://github.com/checkfirst-ltd/nestjs-outlook/commit/617eba5f8fd9cbb565be50e88072646519cdb253))

## [4.6.0](https://github.com/checkfirst-ltd/nestjs-outlook/compare/v4.5.1...v4.6.0) (2025-12-22)


### Features

* use delta sync as a unified source for importing events on first and second connection attempt ([#50](https://github.com/checkfirst-ltd/nestjs-outlook/issues/50)) ([20786f6](https://github.com/checkfirst-ltd/nestjs-outlook/commit/20786f6c30420ac717aae990c69a7770649cb017))

## [4.5.1](https://github.com/checkfirst-ltd/nestjs-outlook/compare/v4.5.0...v4.5.1) (2025-12-22)


### Bug Fixes

* outlook sends notification with empty resources ([#51](https://github.com/checkfirst-ltd/nestjs-outlook/issues/51)) ([1e2aeb4](https://github.com/checkfirst-ltd/nestjs-outlook/commit/1e2aeb43a36d803902bb77d84ee50be7cb985b0f))

## [4.5.0](https://github.com/checkfirst-ltd/nestjs-outlook/compare/v4.4.1...v4.5.0) (2025-12-18)


### Features

* revoke Refresh Token ([#48](https://github.com/checkfirst-ltd/nestjs-outlook/issues/48)) ([4f9e2c6](https://github.com/checkfirst-ltd/nestjs-outlook/commit/4f9e2c6cdc9c5da33c2307c8e686cb1ea12442dc))

## [4.4.1](https://github.com/checkfirst-ltd/nestjs-outlook/compare/v4.4.0...v4.4.1) (2025-12-04)


### Bug Fixes

* create subscription only if there's a READ permission ([#46](https://github.com/checkfirst-ltd/nestjs-outlook/issues/46)) ([ad7fb76](https://github.com/checkfirst-ltd/nestjs-outlook/commit/ad7fb76b0d348381194a6ef2cab1d7a8c503f6d6))

## [4.4.0](https://github.com/checkfirst-ltd/nestjs-outlook/compare/v4.3.0...v4.4.0) (2025-11-19)


### Features

* Migrate publishing to NPM to use OIDC ([e77198d](https://github.com/checkfirst-ltd/nestjs-outlook/commit/e77198d9aec3ca0530cfb154263bcbfc8d94fd99))

## [4.3.0](https://github.com/checkfirst-ltd/nestjs-outlook/compare/v4.2.2...v4.3.0) (2025-11-18)


### Features

* **Calendar:** import calendar events in streamable chunks ([#39](https://github.com/checkfirst-ltd/nestjs-outlook/issues/39)) ([6c3d865](https://github.com/checkfirst-ltd/nestjs-outlook/commit/6c3d865df0590c7a9c434048dc45ce1aec82848d))

## [4.2.2](https://github.com/checkfirst-ltd/nestjs-outlook/compare/v4.2.1...v4.2.2) (2025-10-10)


### Bug Fixes

* **disconnect:** set isActive = false after deleting subscription ([#36](https://github.com/checkfirst-ltd/nestjs-outlook/issues/36)) ([05dfba3](https://github.com/checkfirst-ltd/nestjs-outlook/commit/05dfba33cd0d0bf3075f21ebc78b6928ce8176f1))

## [4.2.1](https://github.com/checkfirst-ltd/nestjs-outlook/compare/v4.2.0...v4.2.1) (2025-09-24)


### Bug Fixes

* external-user-id ([#33](https://github.com/checkfirst-ltd/nestjs-outlook/issues/33)) ([2914f3a](https://github.com/checkfirst-ltd/nestjs-outlook/commit/2914f3a4d4cb29fbcb2d0ca1fc49b8933a076efc))

## [4.2.0](https://github.com/checkfirst-ltd/nestjs-outlook/compare/v4.1.1...v4.2.0) (2025-09-23)


### Features

* add delete event ([#30](https://github.com/checkfirst-ltd/nestjs-outlook/issues/30)) ([7c0527a](https://github.com/checkfirst-ltd/nestjs-outlook/commit/7c0527a33cdb33c1df9ec6b06a3e14e43e41664f))

## [4.1.1](https://github.com/checkfirst-ltd/nestjs-outlook/compare/v4.1.0...v4.1.1) (2025-09-23)


### Bug Fixes

* Fix webhooks sync ([#28](https://github.com/checkfirst-ltd/nestjs-outlook/issues/28)) ([c769ebd](https://github.com/checkfirst-ltd/nestjs-outlook/commit/c769ebd29e629f4b738a48344547745ff203312b))

## [4.1.0](https://github.com/checkfirst-ltd/nestjs-outlook/compare/v4.0.1...v4.1.0) (2025-08-06)


### Features

* added sorting and delta sync ([#23](https://github.com/checkfirst-ltd/nestjs-outlook/issues/23)) ([2eac017](https://github.com/checkfirst-ltd/nestjs-outlook/commit/2eac0176a14784b162fabccbe1462bbac38e9b0f))

## [4.0.1](https://github.com/checkfirst-ltd/nestjs-outlook/compare/v4.0.0...v4.0.1) (2025-05-16)


### Bug Fixes

* Remove unncessary defaults ([dea21c4](https://github.com/checkfirst-ltd/nestjs-outlook/commit/dea21c4e558f12988958bfae1ee577937bdeb558))

## [4.0.0](https://github.com/checkfirst-ltd/nestjs-outlook/compare/v3.1.0...v4.0.0) (2025-05-10)


### ⚠ BREAKING CHANGES

* **auth:** centralize token management with MicrosoftUser entity

### Features

* **auth:** centralize token management with MicrosoftUser entity ([25a538d](https://github.com/checkfirst-ltd/nestjs-outlook/commit/25a538d68b0d6ac522e91e47bcb20d76a8ae8217))

## [3.1.0](https://github.com/checkfirst-ltd/nestjs-outlook/compare/v3.0.0...v3.1.0) (2025-05-10)


### Features

* Implement customizable permission scopes ([05a60b3](https://github.com/checkfirst-ltd/nestjs-outlook/commit/05a60b367d9bd625928e959bac42aa255e335249))


### Bug Fixes

* Make basepath mandatory ([47e4ec9](https://github.com/checkfirst-ltd/nestjs-outlook/commit/47e4ec97fba1d8ac09c88202d474bfac60a99baf))

## [3.0.0](https://github.com/checkfirst-ltd/nestjs-outlook/compare/v2.0.0...v3.0.0) (2025-05-08)


### ⚠ BREAKING CHANGES

* Add support for sending emails

### Features

* Add support for sending emails ([cd66ecd](https://github.com/checkfirst-ltd/nestjs-outlook/commit/cd66ecd3cc05536c54b724c68ec73566b09cc4d0))
* Notify when emails are created/updated/deleted ([eacdfba](https://github.com/checkfirst-ltd/nestjs-outlook/commit/eacdfba7d5667c848a576d043107e2a3962fc121))


### Bug Fixes

* Fix basePath in webhook notifications ([f1b3ff7](https://github.com/checkfirst-ltd/nestjs-outlook/commit/f1b3ff7ae23d60543922911b06eb9c1114273268))

## [2.0.0](https://github.com/checkfirst-ltd/nestjs-outlook/compare/v1.0.0...v2.0.0) (2025-05-05)


### ⚠ BREAKING CHANGES

* **types:** Changed import source for Microsoft Graph types from '@microsoft/microsoft-graph-types' to local types. While functionally identical (re-exports), this change breaks type compatibility for library consumers who directly use these types.

### Features

* **types:** replace Microsoft Graph types with local re-exports ([2110d39](https://github.com/checkfirst-ltd/nestjs-outlook/commit/2110d39d601820bbece827aab262ee157e210f5a))

## 1.0.0 (2025-05-04)


### Features

* initial working module ([64ac682](https://github.com/checkfirst-ltd/nestjs-outlook/commit/64ac6820aa3ba8143bd9919db1d837992e999ec9))
