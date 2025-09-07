<script lang="ts">
    import { lang } from "../Scripts/Language";
import DeleteButton from "./DeleteButton.svelte";

        const {
        sourceShading,
    }: {
        sourceShading: Word.Shading;
    } = $props();
</script>

<label class="flex hcenter gap">
      {lang("Foreground color")}:
      <input
        type="color"
        value={sourceShading.foregroundPatternColor || "#ffffff"}
        onchange={(e) => (sourceShading.foregroundPatternColor = e.currentTarget.value)}
      />
      <DeleteButton callback={(e) => {
        // @ts-ignore
        delete sourceShading.foregroundPatternColor;
        const possibleItem = (e.target as HTMLElement).closest("label")?.querySelector("input[type=color]");
        if (possibleItem) (possibleItem as HTMLInputElement).value = "#ffffff";
        }}></DeleteButton>
    </label><br />
    <label class="flex hcenter gap">
      {lang("Background color")}:
      <input
        type="color"
        value={sourceShading.backgroundPatternColor || "#ffffff"}
        onchange={(e) => (sourceShading.backgroundPatternColor = e.currentTarget.value)}
      />
      <DeleteButton callback={(e) => {
        // @ts-ignore
        delete sourceShading.backgroundPatternColor;
        const possibleItem = (e.target as HTMLElement).closest("label")?.querySelector("input[type=color]");
        if (possibleItem) (possibleItem as HTMLInputElement).value = "#ffffff";
        }}></DeleteButton>
    </label><br />
    <label class="flex hcenter gap">
      {lang("Texture type")}:
      <div class="selectContainer">
        <select bind:value={sourceShading.texture}>
          {#each ["DarkDiagonalDown", "DarkDiagonalUp", "DarkGrid", "DarkHorizontal", "DarkTrellis", "DarkVertical", "LightDiagonalDown", "LightDiagonalUp", "LightGrid", "LightHorizontal", "LightTrellis", "LightVertical", "None", "Percent10", "Percent12Pt5", "Percent15", "Percent20", "Percent25", "Percent30", "Percent35", "Percent37Pt5", "Percent40", "Percent45", "Percent5", "Percent50", "Percent55", "Percent60", "Percent62Pt5", "Percent65", "Percent70", "Percent75", "Percent80", "Percent85", "Percent87Pt5", "Percent90", "Percent95", "Solid", "None"] as option}
            <option value={option}
              >{option.startsWith("Percent")
                ? `${option.replace("Percent", "").replace("Pt", ".")}%`
                : option}</option
            >
          {/each}
        </select>
      </div>
    </label>
