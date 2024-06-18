<template>
  <span class="sort-icon" style="margin-left: 4px">
    <svg
      width="6"
      height="6"
      viewBox="0 0 10 10"
      fill="none"
      xmlns="http://www.w3.org/2000/svg"
    >
      <path d="M5.5 12L0.73686 0.75L10.2631 0.75L5.5 12Z" fill="black" />
    </svg>
  </span>
</template>

<script setup></script>

<style lang="scss">
.sort-icon {
  opacity: 0;
  transition: opacity 0.25s ease-out;
  .--sort-up & {
    opacity: 1;
    svg {
      transform: rotate(180deg);
    }
  }

  .--sort-down & {
    opacity: 1;
    svg {
      transform: rotate(0deg);
    }
  }
}
</style>
