<script lang="ts">
    import { lang } from "../Scripts/Language";

    const {callback}: {callback: (e: Event) => void} = $props();
</script>
<button style="width: fit-content;" title={lang("Discard the changes made to this property")} onclick={callback}>
    -
</button>