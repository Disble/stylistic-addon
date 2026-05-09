# Frontend Guide — User Preferences API

Documento de referencia para el equipo frontend sobre el endpoint de
preferencias globales del usuario.

Este endpoint permite leer y actualizar las instrucciones globales de corrección
que el usuario declara explícitamente para orientar la vigilancia del corrector
estilístico.

## Qué resuelve este endpoint

El perfil del documento aprende patrones a partir del uso y del feedback. Pero
hay defectos que el usuario YA conoce y quiere vigilar desde el primer momento.

Este endpoint existe para eso.

Las instrucciones guardadas aquí:

- son globales por usuario
- no pertenecen a un documento puntual
- solo afectan la corrección activa de texto
- no actualizan el perfil del documento
- no participan en el flujo de feedback

## Dónde vive en la API

- Lectura: `GET /api/user/preferences`
- Escritura: `PUT /api/user/preferences`

Ambos endpoints requieren autenticación con Better Auth bearer token.

```http
Authorization: Bearer <better-auth-session-token>
```

## OpenAPI como fuente navegable

El contrato publicado por el backend puede consultarse en:

- OpenAPI JSON: `/api/openapi.json`
- Swagger UI: `/swagger-ui`

Buscar en OpenAPI por estos operation ids:

- `getUserPreferences`
- `updateUserPreferences`

> Importante: en desarrollo estas rutas están disponibles desde Mastra. En
> producción no hay que asumirlo automáticamente: depende de la configuración de
> build del servidor.

## Casos de uso esperados

### 1. Cargar el formulario de preferencias del usuario

Cuando el usuario abre la pantalla de configuración, el frontend debe pedir el
estado actual para precargar el textarea.

```ts
GET /api/user/preferences
```

Respuesta:

```ts
type UserPreferencesResponse = {
  correctionInstructions: string | null;
  correctionInstructionsMaxLength: 4000;
};
```

Uso esperado en UI:

- si `correctionInstructions` tiene valor, mostrarlo en el campo
- si viene `null`, mostrar estado vacío
- usar `correctionInstructionsMaxLength` para el contador visual y validación de
  UX

### 2. Guardar instrucciones globales de corrección

Cuando el usuario guarda el formulario, el frontend debe enviar el contenido del
campo tal como el usuario lo escribió.

```ts
type UpdateUserPreferencesRequest = {
  correctionInstructions: string | null;
};
```

Ejemplo:

```json
{
  "correctionInstructions": "Vigilá subordinadas demasiado largas, repeticiones léxicas cercanas y adverbios terminados en -mente cuando no aporten precisión."
}
```

### 3. Limpiar las instrucciones guardadas

Si el usuario quiere dejar de usar esta capa global, el frontend debe enviar
`null`.

```json
{
  "correctionInstructions": null
}
```

## Reglas funcionales importantes

- `correctionInstructions` acepta `string | null`
- `null` significa “borrar preferencias globales guardadas”
- el backend recorta espacios al persistir
- si después del trim el contenido queda vacío, el valor persistido efectivo
  queda como `null`
- si el valor es `null`, el workflow de corrección no envía mensaje `system`
  extra al LLM
- longitud máxima: `4000`

## Respuesta exitosa

Tanto `GET` como `PUT` devuelven el mismo shape:

```ts
type UserPreferencesResponse = {
  correctionInstructions: string | null;
  correctionInstructionsMaxLength: 4000;
};
```

Ejemplo con valor guardado:

```json
{
  "correctionInstructions": "Vigilá subordinadas demasiado largas y muletillas repetidas.",
  "correctionInstructionsMaxLength": 4000
}
```

Ejemplo sin valor guardado:

```json
{
  "correctionInstructions": null,
  "correctionInstructionsMaxLength": 4000
}
```

## Errores esperados

### `401 unauthenticated`

La request no tiene una sesión válida.

```json
{
  "error": "unauthenticated"
}
```

### `400 invalid_json_body`

El body no era JSON válido.

```json
{
  "error": "invalid_json_body"
}
```

### `400 invalid_user_preferences_request`

El JSON era válido, pero no respetó el contrato.

```json
{
  "error": "invalid_user_preferences_request",
  "issues": [
    {
      "code": "too_big",
      "maximum": 4000,
      "inclusive": true,
      "path": ["correctionInstructions"],
      "message": "Too big: expected string to have <= 4000 characters"
    }
  ]
}
```

## Cómo se usa desde producto

Pensalo como una capa global de vigilancia.

- El perfil del documento preserva voz, hábitos y patrones aprendidos de ese
  documento.
- `correctionInstructions` agrega focos explícitos declarados por el usuario.

Esto sirve para casos como:

- “siempre revisame subordinadas demasiado largas”
- “prestá atención a muletillas repetidas”
- “marcame abuso de adverbios terminados en -mente”

No sirve para:

- modelar la voz del documento
- aprender estilo automáticamente
- mandar feedback sobre una corrección puntual

## Recomendación de implementación frontend

Flujo sugerido:

1. llamar `GET /api/user/preferences` al abrir configuración
2. poblar textarea y contador con `correctionInstructionsMaxLength`
3. guardar con `PUT /api/user/preferences`
4. usar la respuesta del backend como fuente de verdad final
5. mostrar errores de autenticación o validación con mensajes específicos

## Relación con otros documentos

- Contrato general del flujo estilístico: [`frontend-contract.md`](./frontend-contract.md)
- Autenticación: [`auth.md`](./auth.md)
