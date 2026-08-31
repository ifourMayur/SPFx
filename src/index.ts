/**
 * Solution entry point, exposing the layers in dependency order:
 *
 *   models  ->  config  ->  services  ->  components
 *
 * Components consume these layers and live one folder per component, named after the
 * component itself:
 *
 *   src/webparts/<webPart>/components/<Component>/<Component>.tsx
 *                                                 /I<Component>Props.ts
 *                                                 /<Component>.module.scss
 *
 * Every new component follows that shape, so a component and everything belonging only to
 * it stay together. Web parts and components import from the specific files they need;
 * this barrel documents the public surface of the solution and gives consumers a single
 * import path.
 */
export * from './models';
export * from './config';
export * from './services';
