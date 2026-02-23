<template>
  <div class="app-container">
    <header class="app-header">
      <h1>NAA CAE Outlook Add-in</h1>
      <nav class="app-nav">
        <router-link to="/" class="nav-link">Home</router-link>
        <router-link to="/profile" class="nav-link">Profile</router-link>
      </nav>
    </header>
    <main class="app-main">
      <router-view />
    </main>
  </div>
</template>

<script setup lang="ts">
// Root App component – layout shell
</script>

<style>
/* ---- Reset & global styles ---- */
*,
*::before,
*::after {
  box-sizing: border-box;
  margin: 0;
  padding: 0;
}

body {
  font-family: "Segoe UI", Tahoma, Geneva, Verdana, sans-serif;
  background: #f3f2f1;
  color: #323130;
  font-size: 14px;
}

.app-container {
  max-width: 480px;
  margin: 0 auto;
  padding: 16px;
}

.app-header {
  text-align: center;
  padding: 12px 0 16px;
  border-bottom: 1px solid #edebe9;
  margin-bottom: 16px;
}

.app-header h1 {
  font-size: 18px;
  font-weight: 600;
  color: #0078d4;
}

.app-nav {
  display: flex;
  justify-content: center;
  gap: 16px;
  margin-top: 10px;
}

.nav-link {
  font-size: 13px;
  font-weight: 600;
  color: #0078d4;
  text-decoration: none;
  padding: 4px 12px;
  border-radius: 4px;
  transition: background 0.15s;
}

.nav-link:hover {
  background: #e1effa;
}

.nav-link.router-link-exact-active {
  background: #0078d4;
  color: #fff;
}

.app-main {
  min-height: 300px;
}
</style>
