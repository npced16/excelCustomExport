<template>
  <div class="h-full w-full text-nowrap">
    <h1>Welcome to the Webpage</h1>
    <p>This is a secure webpage for token: {{ token }}</p>
  </div>
</template>

<script setup>
import { ref, onMounted } from 'vue'
import { useRoute } from 'vue-router'
const route = useRoute()
const token = ref('')

onMounted(() => {
  token.value = route.query.token
})
</script>

<style scoped></style>
