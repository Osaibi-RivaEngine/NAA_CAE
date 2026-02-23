<template>
  <div class="home-view">
    <!-- Signed-out state -->
    <section v-if="!isAuthenticated" class="card">
      <h2>Welcome</h2>
      <p>
        Sign in with your Microsoft account to access mailbox features
        via the Nested App Authentication flow.
      </p>
      <button class="btn primary" :disabled="isLoading" @click="login">
        {{ isLoading ? "Signing in…" : "Sign In" }}
      </button>
    </section>

    <!-- Signed-in state -->
    <section v-else class="card">
      <h2>Hello, {{ account?.name ?? account?.username }}!</h2>
      <p class="sub">{{ account?.username }}</p>

      <div class="actions">
        <router-link to="/profile" class="btn primary">
          View Profile (Graph)
        </router-link>
        <button class="btn secondary" @click="readMailbox">
          {{ isLoading ? "Loading…" : "Read Mailbox Item" }}
        </button>
        <button class="btn outline" @click="logout">Sign Out</button>
      </div>

      <div v-if="mailboxInfo" class="result">
        <h3>Current Mail Item</h3>
        <pre>{{ mailboxInfo }}</pre>
      </div>
    </section>

    <!-- Error banner -->
    <div v-if="error" class="error-banner">
      <strong>Error:</strong> {{ error }}
    </div>
  </div>
</template>

<script setup lang="ts">
import { ref } from "vue";
import { useAuth } from "@/composables/useAuth";

const { isAuthenticated, account, isLoading, error, login, logout } = useAuth();

const mailboxInfo = ref<string | null>(null);

/** Use the Office.js mailbox API to read the current item. */
async function readMailbox() {
  try {
    const item = Office.context.mailbox?.item;
    if (!item) {
      mailboxInfo.value = "No mailbox item selected.";
      return;
    }

    mailboxInfo.value = JSON.stringify(
      {
        subject: item.subject,
        from: item.from?.emailAddress,
        dateTimeCreated: item.dateTimeCreated,
        itemType: item.itemType,
      },
      null,
      2
    );
  } catch (e: unknown) {
    mailboxInfo.value = `Error: ${(e as Error).message}`;
  }
}
</script>

<style scoped>
.card {
  background: #fff;
  border-radius: 8px;
  padding: 20px;
  box-shadow: 0 1px 4px rgba(0, 0, 0, 0.08);
  margin-bottom: 16px;
}

.card h2 {
  font-size: 16px;
  margin-bottom: 8px;
}

.card .sub {
  color: #605e5c;
  font-size: 13px;
  margin-bottom: 12px;
}

.actions {
  display: flex;
  flex-direction: column;
  gap: 8px;
  margin-top: 12px;
}

.btn {
  display: inline-block;
  text-align: center;
  padding: 10px 16px;
  border-radius: 4px;
  font-size: 14px;
  font-weight: 600;
  cursor: pointer;
  border: none;
  text-decoration: none;
  transition: background 0.15s;
}

.btn.primary {
  background: #0078d4;
  color: #fff;
}
.btn.primary:hover {
  background: #106ebe;
}

.btn.secondary {
  background: #edebe9;
  color: #323130;
}
.btn.secondary:hover {
  background: #d2d0ce;
}

.btn.outline {
  background: transparent;
  border: 1px solid #8a8886;
  color: #323130;
}
.btn.outline:hover {
  background: #f3f2f1;
}

.btn:disabled {
  opacity: 0.6;
  cursor: not-allowed;
}

.result {
  margin-top: 16px;
  background: #faf9f8;
  padding: 12px;
  border-radius: 4px;
}

.result h3 {
  font-size: 14px;
  margin-bottom: 8px;
}

.result pre {
  font-size: 12px;
  white-space: pre-wrap;
  word-break: break-word;
}

.error-banner {
  margin-top: 12px;
  padding: 10px 14px;
  background: #fde7e9;
  color: #a80000;
  border-radius: 4px;
  font-size: 13px;
}
</style>
