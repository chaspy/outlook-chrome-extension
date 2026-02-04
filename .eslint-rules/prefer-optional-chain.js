"use strict";

const isSameReference = (left, right) => {
  if (!left || !right) return false;
  if (left.type === "Identifier" && right.type === "Identifier") {
    return left.name === right.name;
  }
  if (left.type === "MemberExpression" && right.type === "MemberExpression") {
    if (left.computed || right.computed) return false;
    if (!left.property || !right.property) return false;
    if (left.property.type !== "Identifier" || right.property.type !== "Identifier") return false;
    if (left.property.name !== right.property.name) return false;
    return isSameReference(left.object, right.object);
  }
  return false;
};

const getMatchTarget = (right) => {
  if (!right) return null;
  if (right.type === "MemberExpression") return right.object;
  if (right.type === "CallExpression") return right.callee;
  return null;
};

module.exports = {
  meta: {
    type: "suggestion",
    docs: {
      description: "Prefer optional chaining when accessing properties after null checks"
    },
    schema: [],
    messages: {
      preferOptionalChain: "Prefer using optional chaining for this access."
    }
  },
  create(context) {
    return {
      LogicalExpression(node) {
        if (node.operator !== "&&") return;
        const target = getMatchTarget(node.right);
        if (!target) return;
        if (isSameReference(node.left, target)) {
          context.report({ node, messageId: "preferOptionalChain" });
        }
      }
    };
  }
};
