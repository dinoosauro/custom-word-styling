<script lang="ts">
    import { cubicInOut } from "svelte/easing";
    import { slide } from "svelte/transition";
    let { children, secondCard }: {
        children: any,
        secondCard?: boolean | undefined
    }= $props();
</script>

<div
    class={`card${secondCard ? " secondCard" : ""}`}
    in:slide={{ duration: 400, easing: cubicInOut }}
    out:slide={{ duration: 400, easing: cubicInOut }}
>
    {@render children?.()}
</div>
