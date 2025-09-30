<script lang="ts">
    import { AvailableShapes } from "../Scripts/HelperType";
    import { lang } from "../Scripts/Language";
    import CreateRoundedShape from "../Scripts/RoundedShape";
    import Card from "./Card.svelte";
    import GradientShapePicker from "./GradientShapePicker.svelte";

    let width = 200;
    let height = 200;
    let borderRadius = 0.2;
    /**
     * The fill type. Can be `color` for a plain color, `gradient` for a linear gradient, or `image` for a background image.
     */
    let fillType = $state("color");
    /**
     * The border color type. Can be `none` (empty), `color` for a plain color, or `gradient` for a linear gradient.
     */
    let borderFillType = $state("none");
    /**
     * The selected color, in the "Plain color" mode
     */
    let selectedColor = "#cccccc";
    /**
     * Transparency of the plain color added inside the shape
     */
    let colorTransparency = 0;
    /**
     * The selected border color, in the "Plain color" mode
     */
    let selectedBorderColor = "#cccccc";
    /**
     * Transparency of the plain border color
     */
    let borderColorTransparency = 0;

    // Gradient variables
    /**
     * Gradient of the content inside the rounded rectangle
     */
    let insideGradient = {
        colors: [],
        gradientDegrees: 180,
        currentColor: "#cccccc",
        currentPercentage: 0,
        currentTransparency: 0,
    };
    /**
     * Border gradient
     */
    let outsideGradient = {
        colors: [],
        gradientDegrees: 180,
        currentColor: "#cccccc",
        currentPercentage: 0,
        currentTransparency: 0,
    };
    // Image variables
    /**
     * An array that contains [the base64 of the image, the file name, the image width, the image height]
     */
    let selectedImage = $state(["", "", 1000, 1000]);
    let saveAsPng = false;
    let quality = 0.8;
    let scaleType: "none" | "keepWidth" | "keepHeight" = "none";
    let borderSize = 1;
    /**
     * Re-encode the passed Blob, and get its base64
     * @param source the Blob with the image content
     * @param name the file name
     */
    function updateSelectedImage(source: Blob, name: string) {
        const img = new Image();
        img.onload = () => {
            const canvas = Object.assign(document.createElement("canvas"), {
                width: img.naturalWidth,
                height: img.naturalHeight,
            });
            canvas
                .getContext("2d")
                ?.drawImage(img, 0, 0, canvas.width, canvas.height);
            const url = canvas.toDataURL(
                saveAsPng ? "image/png" : "image/jpeg",
                quality,
            );
            selectedImage = [
                url.substring(url.indexOf(",") + 1),
                name,
                img.naturalWidth,
                img.naturalHeight,
            ];
        };
        img.src = URL.createObjectURL(source);
    }
    /**
     * The type of the border line (ex: single, double etc.)
     */
    let borderType = "sng";
    /**
     * The style of the border line (ex: normal, dashed, dotted etc.).
     * It usually contains two values, separated by a space, so that two different properties can be updated.
     */
    let borderLineStyle = "normal";
    /**
     * The selected shape type
     */
    let outputShape = "RoundRectangle";
</script>

<h2>{lang("Custom shape")}:</h2>
<label class="flex hcenter gap">
    {lang("Shape type")}:
    <div class="selectContainer">
        <select bind:value={outputShape}>
            {#each AvailableShapes as option}
            <option value={option}>{option}</option>
            {/each}
        </select>
    </div>
</label><br>
<label class="flex hcenter gap">
    {lang("Width (in points):")}
    <input type="number" min="0" bind:value={width} />
</label><br />
<label class="flex hcenter gap">
    {lang("Height (in points):")}
    <input type="number" min="0" bind:value={height} />
</label><br />
<label class="flex hcenter gap">
    {lang("Border radius, from 0 (square) to 0.5 (circle):")}
    <input
        type="number"
        min="0"
        max="0.5"
        step="0.001"
        bind:value={borderRadius}
    />
</label><br />
<label class="flex hcenter gap">
    {lang("Fill")}:
    <div class="selectContainer">
        <select bind:value={fillType}>
            <option value="none">{lang("None")}</option>
            <option value="color">{lang("Single color")}</option>
            <option value="gradient">{lang("Gradient")}</option>
            <option value="image">{lang("Image")}</option>
        </select>
    </div>
</label><br />
{#if fillType === "color"}
    <Card secondCard={true}>
        <label class="flex hcenter gap">
            {lang("Select the color")}:
            <input bind:value={selectedColor} type="color" />
        </label><br />
        <label class="flex hcenter gap">
            {lang("Transparency")}:
            <input
                type="range"
                min="0"
                max="1"
                step="0.001"
                bind:value={colorTransparency}
            />
        </label>
    </Card>
{:else if fillType === "gradient"}
    <Card secondCard={true}>
        <GradientShapePicker data={insideGradient}></GradientShapePicker>
    </Card>
{:else}
    <Card secondCard={true}>
        <label class="flex hcenter gap">
            <input type="checkbox" bind:checked={saveAsPng} />{lang(
                "Add a PNG instead of a JPEG",
            )}
        </label><br />
        <label class="flex hcenter gap">
            {lang("Output image quality (for JPEG)")}:
            <input
                type="range"
                min="0"
                max="1"
                step="0.01"
                bind:value={quality}
            />
        </label><br />
        <label class="flex hcenter gap">
            {lang("Keep aspect ratio by resizing")}
            <div class="selectContainer">
                <select bind:value={scaleType}>
                    <option value="keepWidth">{lang("the height")}</option>
                    <option value="keepHeight">{lang("the width")}</option>
                    <option value="none">{lang("none of the above")}</option>
                </select>
            </div>
        </label><br />
        <div class="flex gap">
            <button
                onclick={() => {
                    const input = Object.assign(
                        document.createElement("input"),
                        {
                            type: "file",
                            accept: "image/*",
                            onchange: () => {
                                input.files &&
                                    updateSelectedImage(
                                        input.files[0],
                                        input.files[0].name,
                                    );
                            },
                        },
                    );
                    input.click();
                }}>{lang("Pick image")}</button
            >
            <button
                onclick={async () => {
                    const clipboard = await navigator.clipboard.read();
                    for (const item of clipboard) {
                        const imageType = item.types.find((i) =>
                            i.startsWith("image/"),
                        );
                        if (!imageType) continue;
                        const imageItem = await item.getType(imageType);
                        updateSelectedImage(
                            imageItem,
                            `Clipboard-${new Date().toLocaleString()}`,
                        );
                        break;
                    }
                }}>{lang("Copy from clipboard")}</button
            >
        </div>
        <br />
        <i
            >{selectedImage[0] === ""
                ? lang("Select an image")
                : `${lang("Selected image")}: ${selectedImage[1]}`}</i
        >
    </Card>
{/if}
<br />
<label class="flex hcenter gap">
    {lang("Border color:")}
    <div class="selectContainer">
        <select bind:value={borderFillType}>
            <option value="none">{lang("None")}</option>
            <option value="color">{lang("Plain color")}</option>
            <option value="gradient">{lang("Linear gradient")}</option>
        </select>
    </div>
</label><br />
{#if borderFillType !== "none"}
    <label class="flex hcenter gap">
        {lang("Border size (in points)")}:
        <input type="number" bind:value={borderSize} />
    </label><br />
    <Card secondCard={true}>
        {#if borderFillType === "color"}
            <label class="flex hcenter gap">
                {lang("Border color:")}
                <input type="color" bind:value={selectedBorderColor} />
            </label><br />
            <label class="flex hcenter gap">
                {lang("Transparency:")}
                <input
                    type="range"
                    min="0"
                    max="1"
                    step="0.001"
                    bind:value={borderColorTransparency}
                />
            </label>
        {:else}
            <GradientShapePicker data={outsideGradient}></GradientShapePicker>
        {/if}
    </Card><br />
    <Card secondCard={true}>
        <h3>{lang("Advanced border settings")}:</h3>
        <label class="flex hcenter gap">
            {lang("Line type")}:
            <div class="selectContainer">
                <select bind:value={borderType}>
                    <option value="sng">{lang("Single")}</option>
                    <option value="dbl">{lang("Double")}</option>
                    <option value="thickThin"
                        >{lang("Double (thick and thin)")}</option
                    >
                    <option value="thinThick"
                        >{lang("Double (thin and tick)")}</option
                    >
                    <option value="tri">{lang("Triple")}</option>
                </select>
            </div>
        </label><br />
        <label class="flex hcenter gap">
            {lang("Line style")}:
            <div class="selectContainer">
                <select bind:value={borderLineStyle}>
                    <option value="normal">{lang("Normal")}</option>
                    <option value="sysDot 1 1">{lang("Dotted")}</option>
                    <option value="sysDash 3 1">{lang("Dashed")}</option>
                    <option value="dash dash">{lang("Long dashes")}</option>
                    <option value="dashDot dashDot"
                        >{lang("Dashes with dots")}</option
                    >
                    <option value="lgDash longDash"
                        >{lang("Really long dashes")}</option
                    >
                    <option value="lgDashDot longDashDot"
                        >{lang("Really long dashes with dots")}</option
                    >
                    <option value="lgDashDotDot longDashDotDot"
                        >{lang("Really long dashes with double dots")}</option
                    >
                </select>
            </div>
        </label>
    </Card><br />
{/if}
<button
    onclick={() => {
        CreateRoundedShape({
            color: selectedColor,
            gradient:
                fillType === "gradient"
                    ? {
                          colors: insideGradient.colors,
                          angle: insideGradient.gradientDegrees,
                      }
                    : undefined,
            image:
                fillType === "image"
                    ? {
                          isPng: saveAsPng,
                          base64: selectedImage[0] as string,
                          width: selectedImage[2] as number,
                          height: selectedImage[3] as number,
                          scaleType,
                      }
                    : undefined,
            width,
            height,
            borderRadius,
            border:
                borderFillType === "none"
                    ? undefined
                    : {
                          plainColor:
                              borderFillType === "color"
                                  ? selectedBorderColor
                                  : undefined,
                          plainColorTransparency: borderColorTransparency,
                          gradient:
                              borderFillType === "gradient"
                                  ? {
                                        angle: outsideGradient.gradientDegrees,
                                        colors: outsideGradient.colors,
                                    }
                                  : undefined,
                          size: borderSize,
                          borderType,
                          borderLineStyle,
                      },
            plainColorSettings: {
                transparency: colorTransparency,
                deleteBackground: fillType === "none",
            },
            shapeType: outputShape
        });
    }}>{lang("Add shape")}</button
>
