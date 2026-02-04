"use strict";

const getLiteralValue = (node) => {
  if (!node) return null;
  if (node.type === "Literal") {
    return { type: typeof node.value, value: node.value };
  }
  if (
    node.type === "UnaryExpression" &&
    node.operator === "-" &&
    node.argument &&
    node.argument.type === "Literal" &&
    typeof node.argument.value === "number"
  ) {
    return { type: "number", value: -node.argument.value };
  }
  return null;
};

module.exports = {
  meta: {
    type: "problem",
    docs: {
      description: "Disallow functions that always return the same constant value"
    },
    schema: [],
    messages: {
      constantReturn: "Refactor this function to not always return the same value."
    }
  },
  create(context) {
    const stack = [];

    const enterFunction = (node) => {
      stack.push({ node, returns: [], totalReturns: 0, hasNonLiteral: false });
    };

    const exitFunction = () => {
      const current = stack.pop();
      if (!current || current.totalReturns < 2) return;
      if (current.hasNonLiteral) return;
      const [first, ...rest] = current.returns;
      if (!first) return;
      const same = rest.every(
        (value) => value && value.type === first.type && value.value === first.value
      );
      if (same) {
        context.report({ node: current.node, messageId: "constantReturn" });
      }
    };

    return {
      FunctionDeclaration: enterFunction,
      "FunctionDeclaration:exit": exitFunction,
      FunctionExpression: enterFunction,
      "FunctionExpression:exit": exitFunction,
      ArrowFunctionExpression: enterFunction,
      "ArrowFunctionExpression:exit": exitFunction,
      ReturnStatement(node) {
        if (!stack.length) return;
        const current = stack[stack.length - 1];
        current.totalReturns += 1;
        const value = getLiteralValue(node.argument);
        if (!value) {
          current.hasNonLiteral = true;
          return;
        }
        current.returns.push(value);
      }
    };
  }
};
