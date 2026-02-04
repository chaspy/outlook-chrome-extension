module.exports = {
  env: {
    browser: true,
    es2021: true
  },
  parserOptions: {
    ecmaVersion: 2021,
    sourceType: "script"
  },
  plugins: ["sonarjs", "unicorn"],
  extends: ["eslint:recommended"],
  globals: {
    chrome: "readonly"
  },
  rules: {
    "max-nested-callbacks": ["error", 4],
    "no-constant-return": "error",
    "no-empty": ["error", { allowEmptyCatch: false }],
    "no-restricted-properties": [
      "error",
      {
        object: "document",
        property: "execCommand",
        message: "document.execCommand is deprecated."
      },
      {
        object: "document",
        property: "queryCommandSupported",
        message: "document.queryCommandSupported is deprecated."
      },
      {
        property: "insertAdjacentElement",
        message: "Prefer Element#after() to insertAdjacentElement('afterend', ...)."
      }
    ],
    "no-unused-vars": ["error", { argsIgnorePattern: "^_" }],
    "prefer-global-this": "error",
    "prefer-optional-chain": "error",
    "sonarjs/cognitive-complexity": ["error", 15],
    "unicorn/prefer-code-point": "error",
    "unicorn/prefer-dom-node-dataset": "error",
    "unicorn/prefer-string-replace-all": "error"
  }
};
