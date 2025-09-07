<script lang="ts">
    import type { HelperType } from "../Scripts/HelperType";
    import { lang } from "../Scripts/Language";
    import FontChange from "./FontChange.svelte";
    import HelperDialogs from "./HelperDialogs.svelte";

    const {
        sourceList,
    }: {
        sourceList: Word.ListTemplate;
    } = $props();
    let listItems = $state(0);
    let helperProp: HelperType | undefined = $state(undefined);
</script>

<label class="flex hcenter gap">
    <input type="checkbox" bind:checked={sourceList.outlineNumbered} />{lang("Is an ordered (numeric) list")}
</label><br /><br />
<div class="secondCard">
    <div class="selectContainer">
        <select bind:value={listItems}>
            {#each sourceList.listLevels.items as item, i}
                <option value={i}>{lang("Position number")} {i + 1}</option>
            {/each}
        </select>
    </div>
    <br />
    <label class="flex hcenter gap">
        {lang("Alignment")}:
        <div class="selectContainer">
            <select
                bind:value={sourceList.listLevels.items[listItems].alignment}
            >
                {#each ["Unknown", "Left", "Centered", "Right", "Justified"] as option}
                    <option value={option}>{lang(option)}</option>
                {/each}
            </select>
        </div>
    </label><br />
    <label class="flex hcenter gap">
        {lang("An item of list level")}
        <input
            type="number"
            style="width: 60px;"
            bind:value={sourceList.listLevels.items[listItems].resetOnHigher}
        />
        {lang("should appear before restarting counting from 1")}
    </label><br />
    <div class="card">
        <h3>{lang("Blank spaces")}:</h3>
        <label class="flex hcenter gap">
            {lang("Blank space (in points) at the left of the list")}
            <input
                type="number"
                bind:value={
                    sourceList.listLevels.items[listItems].numberPosition
                }
            />
        </label><br />
        <label class="flex hcenter gap">
            <span class="help" onclick={() => (helperProp = "ListPosition")}>{lang("Tab position")}:</span>
            <input
                type="number"
                bind:value={sourceList.listLevels.items[listItems].tabPosition}
            />
        </label><br />
        <label class="flex hcenter gap">
            <span class="help" onclick={() => (helperProp = "ListSecondLine")}>{lang("Position (in points) of the second line of text")}</span>
            <input
                type="number"
                bind:value={sourceList.listLevels.items[listItems].textPosition}
            />
        </label><br />
        <label class="flex hcenter gap">
            {lang("After the number/symbol, add a:")}
            <div class="selectContainer">
                <select
                    bind:value={
                        sourceList.listLevels.items[listItems].trailingCharacter
                    }
                >
                    {#each [["TrailingTab", "tabulation"], ["TrailingSpace", "space"], ["TrailingNone", "nothing"]] as [key, display]}
                        <option value={key}>{lang(display)}</option>
                    {/each}
                </select>
            </div>
        </label>
    </div>
    <br />
    <div class="card">
        <h3>{lang("Number-specific list options")}:</h3>
        <label class="flex hcenter gap">
            <span class="help" onclick={() => (helperProp = "ListNumberFormat")}>{lang("Number format")}:</span>
            <input
                type="text"
                bind:value={sourceList.listLevels.items[listItems].numberFormat}
            />
        </label><br />
        <label class="flex hcenter gap">
            {lang("Number style")}: <div class="selectContainer">
                <select
                    bind:value={
                        sourceList.listLevels.items[listItems].numberStyle
                    }
                >
                    {#each ["None", "Arabic", "UpperRoman", "LowerRoman", "UpperLetter", "LowerLetter", "Ordinal", "CardinalText", "OrdinalText", "Kanji", "KanjiDigit", "AiueoHalfWidth", "IrohaHalfWidth", "ArabicFullWidth", "KanjiTraditional", "KanjiTraditional2", "NumberInCircle", "Aiueo", "Iroha", "ArabicLZ", "Bullet", "Ganada", "Chosung", "GBNum1", "GBNum2", "GBNum3", "GBNum4", "Zodiac1", "Zodiac2", "Zodiac3", "TradChinNum1", "TradChinNum2", "TradChinNum3", "TradChinNum4", "SimpChinNum1", "SimpChinNum2", "SimpChinNum3", "SimpChinNum4", "HanjaRead", "HanjaReadDigit", "Hangul", "Hanja", "Hebrew1", "Arabic1", "Hebrew2", "Arabic2", "HindiLetter1", "HindiLetter2", "HindiArabic", "HindiCardinalText", "ThaiLetter", "ThaiArabic", "ThaiCardinalText", "VietCardinalText", "LowercaseRussian", "UppercaseRussian", "LowercaseGreek", "UppercaseGreek", "ArabicLZ2", "ArabicLZ3", "ArabicLZ4", "LowercaseTurkish", "UppercaseTurkish", "LowercaseBulgarian", "UppercaseBulgarian", "PictureBullet", "Legal", "LegalLZ"] as obj}
                        <option value={obj}>{obj}</option>
                    {/each}
                </select>
            </div>
        </label><br />
        <label class="flex hcenter gap">
            {lang("Starting number of the list")}:
            <input
                type="number"
                bind:value={sourceList.listLevels.items[listItems].startAt}
            />
        </label>
    </div>
    <br />
    <div class="card">
        <h3>{lang("Font options (probably won't work)")}:</h3>
        {#if sourceList.listLevels.items[listItems].font}
            <FontChange sourceFont={sourceList.listLevels.items[listItems].font}
            ></FontChange>
        {/if}
    </div>
</div>

{#if helperProp}
<HelperDialogs helperType={helperProp} callback={() => (helperProp = undefined)}></HelperDialogs>
{/if}