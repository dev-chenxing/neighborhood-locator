import { defineConfig, presetWind4, presetWebFonts } from "unocss";

export default defineConfig({
  presets: [
    presetWind4(),
    presetWebFonts({
      provider: "bunny",
      fonts: {
        // "noto-sans": "Noto Sans SC:400",
        "noto-serif": "Noto Serif SC:700",
      },
    }),
  ],
  shortcuts: {
    "font-xiaobiaosong": "font-['Times_New_Roman','方正小标宋\\_GBK','Noto_Serif_SC']",
    "font-fangsong": "font-['Times_New_Roman','仿宋\\_GB2312','Noto_Serif_SC']",
  },
});
