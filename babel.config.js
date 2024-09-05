import transformRuntime from "@babel/plugin-transform-runtime";
import presetEnv from "@babel/preset-env";
import presetTypescript from "@babel/preset-typescript";

export default {
  presets: [
    [
      presetEnv,
      {
        targets: {
          esmodules: true,
        },
      },
    ],
    [presetTypescript],
  ],
  plugins: [transformRuntime],
};
