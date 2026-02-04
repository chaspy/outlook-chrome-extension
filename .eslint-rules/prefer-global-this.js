"use strict";

const isDeclaration = (node) => {
  if (!node || !node.parent) return false;
  const parent = node.parent;
  if (parent.type === "VariableDeclarator" && parent.id === node) return true;
  if ((parent.type === "FunctionDeclaration" || parent.type === "FunctionExpression") && parent.id === node) {
    return true;
  }
  if ((parent.type === "ClassDeclaration" || parent.type === "ClassExpression") && parent.id === node) {
    return true;
  }
  if (parent.type === "CatchClause" && parent.param === node) return true;
  if ((parent.type === "ImportSpecifier" || parent.type === "ImportDefaultSpecifier" || parent.type === "ImportNamespaceSpecifier") && parent.local === node) {
    return true;
  }
  if (parent.type === "Property" && parent.key === node) return true;
  if (parent.type === "MethodDefinition" && parent.key === node) return true;
  if (parent.type === "MemberExpression" && parent.property === node && !parent.computed) return true;
  return false;
};

module.exports = {
  meta: {
    type: "suggestion",
    docs: {
      description: "Prefer globalThis over window"
    },
    schema: [],
    messages: {
      preferGlobalThis: "Prefer using globalThis over window."
    }
  },
  create(context) {
    return {
      Identifier(node) {
        if (node.name !== "window") return;
        if (isDeclaration(node)) return;
        context.report({ node, messageId: "preferGlobalThis" });
      }
    };
  }
};
