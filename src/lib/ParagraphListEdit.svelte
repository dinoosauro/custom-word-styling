<script lang="ts">
    import { slide } from "svelte/transition";
    import { lang } from "../Scripts/Language";
    import { cubicInOut } from "svelte/easing";
    const spinner = document.createElement("div");
    spinner.classList.add("spinner");
    let info = $state({
        alignment: "Don't change",
        bulletInfo: {
            changeBulletType: false,
            bulletType: "Solid",
            customCharacter: "*",
            customFont: "Arial",
        },
        indent: {
            changeIndent: false,
            newIndent: 0,
            indentBeforePoint: 0
        },
        numberInfo: {
            changeNumber: false,
            numberType: "Arabic",
            changeSyntax: false,
            numberSyntax: "$1"
        },
        bulletImage: "",
        startFrom: {
            change: false,
            number: 1
        }
    });
    let selectedItem = 1;
</script>

<p>
    {lang("Here you can change the selected lists. Note that you need to save the settings in this section every time you edit a level, not at the end with the \"Apply edits\" button outside this card.")}
</p>
<br />
<label class="flex hcenter gap">
    {lang("Edit level")} <input type="number" min="1" bind:value={selectedItem}>
</label><br />
<div class="secondCard">
    <label class="flex hcenter gap">
        {lang("Alignment")}: <div class="selectContainer"><select bind:value={info.alignment}>
            {#each ["Unknown", "Left", "Centered", "Right", "Justified", "Don't change"] as option}
                <option value={option}>{lang(option)}</option>
            {/each}
        </select></div>
    </label><br />
    <label class="flex hcenter gap">
        <input type="checkbox" bind:checked={info.bulletInfo.changeBulletType} />{lang("Change bullet type")}
    </label>
    {#if info.bulletInfo.changeBulletType}
        <div     in:slide={{ duration: 400, easing: cubicInOut }}
    out:slide={{ duration: 400, easing: cubicInOut }}
>
            <br />
            <div class="card">
                <label class="flex hcenter gap">
                    {lang("Bullet type")}: <div class="selectContainer"><select bind:value={info.bulletInfo.bulletType}>
                        {#each ["Custom", "Solid", "Hollow", "Square", "Diamonds", "Arrow", "Checkmark"] as item}
                            <option value={item}>{item}</option>
                        {/each}
                    </select></div>
                </label>
                {#if info.bulletInfo.bulletType === "Custom"}
                    <div in:slide={{ duration: 400, easing: cubicInOut }}
    out:slide={{ duration: 400, easing: cubicInOut }}>
                        <br />
                        <label class="flex hcenter gap">
                            {lang("Custom character")}: <input
                                type="text"
                                maxlength="1"
                                bind:value={info.bulletInfo.customCharacter}
                            />
                        </label><br />
                        <label class="flex hcenter gap">
                            {lang("Font of the custom character")}: <input
                                type="text"
                                bind:value={info.bulletInfo.customFont}
                            />
                        </label>
                    </div>
                {/if}
            </div>
        </div>
    {/if}<br>
    <label class="flex hcenter gap">
        <input bind:checked={info.indent.changeIndent} type="checkbox">{lang("Change indentation")}
    </label>
    {#if info.indent.changeIndent}
    <div     in:slide={{ duration: 400, easing: cubicInOut }}
    out:slide={{ duration: 400, easing: cubicInOut }}
>
        <br>
        <div class="card">
            <label class="flex hcenter gap">
                {lang("Text indentation")}: <input type="number" bind:value={info.indent.newIndent}>
            </label><br>
            <label class="flex hcenter gap">
                {lang("Space left at the left of the bullet point")}: <input type="number" bind:value={info.indent.indentBeforePoint}>
            </label>
        </div>
    </div>
    {/if}<br>
    <label class="flex hcenter gap">
        <input type="checkbox" bind:checked={info.numberInfo.changeNumber}>{lang("Change number type")}
    </label>
    {#if info.numberInfo.changeNumber}
    <div     in:slide={{ duration: 400, easing: cubicInOut }}
    out:slide={{ duration: 400, easing: cubicInOut }}
>
        <br>
        <div class="card">
            <label class="flex hcenter gap">
                {lang("Number type")}: <div class="select"><select bind:value={info.numberInfo.numberType}>
                    {#each ["None","Arabic","UpperRoman","LowerRoman","UpperLetter","LowerLetter"] as option}
                    <option value={option}>{option}</option>
                    {/each}
                </select></div>
            </label><br>
            <label class="flex hcenter gap">
                <input type="checkbox" bind:checked={info.numberInfo.changeSyntax}>{lang("Change the number format")}
            </label>
            {#if info.numberInfo.changeSyntax}
            <div in:slide={{ duration: 400, easing: cubicInOut }} out:slide={{ duration: 400, easing: cubicInOut }}>
                <br>
                <div class="secondCard">
                <p>{lang("Write here the number format. Note that:")}</p>
                <ul>
                        <li>{lang("You can reference the values of the current level by writing")} <code>${selectedItem}$</code></li>
                        <li>{lang("You can reference previous levels by writing the dollar sign followed by the list level, and then another dollar sign")}</li>
                        <li>{lang("If you need to use the dollar sign everywhere else, write")} <code>$dollar</code> {lang("(without the closing dollar sign!)")}</li>
                    </ul>
                <input type="text" bind:value={info.numberInfo.numberSyntax}>
            </div>
            </div>
            {/if}
        </div>
    </div>
    {/if}<br>
    <label class="flex hcenter gap">
        <input type="checkbox" checked={info.bulletImage !== ""} onchange={(e) => {
            const checked = (e.target as HTMLInputElement).checked;
            info.bulletImage = "";
            (e.target as HTMLInputElement).checked = false;
            if (checked) {
                const input = Object.assign(document.createElement("input"), {
                    accept: "image/*",
                    type: "file",
                    onchange: () => {
                    if (input.files) {
                        let reader = new FileReader();
                        reader.onload = () => {
                            info.bulletImage = (reader.result as string).substring((reader.result as string).indexOf(",") + 1); // Let's remove the "data,base64" string
                        }
                        reader.readAsDataURL(input.files[0]);
                    }
                    }
                });
                input.click();
            }
        }}>{lang("Use an image for the current bullet point")}
    </label><br>
    <label class="flex hcenter gap">
        <input type="checkbox" bind:checked={info.startFrom.change}>{lang("Change the starting number of the list")}
    </label>
    {#if info.startFrom.change}
    <div     in:slide={{ duration: 400, easing: cubicInOut }}
    out:slide={{ duration: 400, easing: cubicInOut }}
>
        <br>
        <div class="card">
            <label class="flex hcenter gap">
                {lang("Start from")}: <input type="number" bind:value={info.startFrom.number}>
            </label>
        </div>
    </div>
    {/if}<br>
    <button onclick={() => {
        document.body.append(spinner);
        setTimeout(async () => {
            await Word.run(async (ctx) => {
                const paragraph = ctx.document.getSelection().load();
                await ctx.sync();
                const lists = paragraph.lists.load();
                await ctx.sync();
                if (lists?.items?.length > 0) {
                    for (const list of lists.items) {
                        if (info.alignment !== "Don't change") list.setLevelAlignment(+selectedItem - 1, info.alignment as "Left");
                        if (info.bulletImage !== "") list.setLevelPicture(+selectedItem - 1, info.bulletImage);
                        if (info.bulletInfo.changeBulletType) list.setLevelBullet(+selectedItem - 1, info.bulletInfo.bulletType as "Custom", info.bulletInfo.bulletType === "Custom" ? info.bulletInfo.customCharacter.charCodeAt(0) : undefined, info.bulletInfo.bulletType === "Custom" ? info.bulletInfo.customFont : undefined);
                        if (info.indent.changeIndent) list.setLevelIndents(+selectedItem - 1, info.indent.newIndent, info.indent.indentBeforePoint);
                        if (info.numberInfo.changeNumber) {
                            if (info.numberInfo.changeSyntax) { // We need to update the number formatting: the numbers between dollar signs are a reference to the previous/current levels, while $dollar is the dollar sign escape code
                                const splitText = info.numberInfo.numberSyntax.split(/\$(?!dollar)/);
                                /**
                                 * Get either the number or the string for Word's level numbering array
                                 * @param str the source string
                                 */
                                function intelligentParse(str: string) {
                                    if (str === "") return undefined;
                                    if (!isNaN(+str)) return +str - 1;
                                    return str;
                                }
                                let outputArr = splitText.map(i => {
                                    let text = intelligentParse(i);
                                    if (typeof text === "string") text = text.replaceAll("$dollar", "$");
                                    return text;
                                }).filter(i => typeof i !== "undefined");
                                list.setLevelNumbering(+selectedItem - 1, info.numberInfo.numberType as "Arabic", outputArr);
                            } else list.setLevelNumbering(+selectedItem - 1, info.numberInfo.numberType as "Arabic");
                        }
                        if (info.startFrom.change) list.setLevelStartingNumber(+selectedItem - 1, info.startFrom.number);
                    }
                }
                await ctx.sync();
                spinner.remove();
            })
        }, 5)
    }}>{lang("Save")}</button>
</div>
