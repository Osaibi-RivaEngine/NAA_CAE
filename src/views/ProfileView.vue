<template>
  <div class="profile-view">
    <router-link to="/" class="back-link">&larr; Back</router-link>

    <section v-if="isLoading" class="card">
      <p>Loading profile from Microsoft Graph…</p>
    </section>

    <section v-else-if="profile" class="card">
      <h2>Profile</h2>
      <table class="profile-table">
        <tbody>
          <tr>
            <td class="label">Display Name</td>
            <td>{{ profile.displayName }}</td>
          </tr>
          <tr>
            <td class="label">Email</td>
            <td>{{ profile.mail ?? profile.userPrincipalName }}</td>
          </tr>
          <tr>
            <td class="label">Job Title</td>
            <td>{{ profile.jobTitle ?? "—" }}</td>
          </tr>
          <tr>
            <td class="label">Office</td>
            <td>{{ profile.officeLocation ?? "—" }}</td>
          </tr>
          <tr>
            <td class="label">Phone</td>
            <td>{{ profile.businessPhones?.join(", ") || "—" }}</td>
          </tr>
        </tbody>
      </table>

      <h3 style="margin-top: 20px">Raw Graph Response</h3>
      <pre class="raw">{{ JSON.stringify(profile, null, 2) }}</pre>
    </section>

    <div v-if="error" class="error-banner">
      <strong>Error:</strong> {{ error }}
      <p v-if="isClaimsError" class="hint">
        A claims challenge was detected. The app automatically retried
        with the new claims. If you still see this error, a Conditional
        Access policy may be blocking access.
      </p>
    </div>
  </div>
</template>

<script setup lang="ts">
import { ref, computed, onMounted } from "vue";
import { useAuth } from "@/composables/useAuth";

interface GraphProfile {
  displayName?: string;
  mail?: string;
  userPrincipalName?: string;
  jobTitle?: string;
  officeLocation?: string;
  businessPhones?: string[];
  [key: string]: unknown;
}

const { isLoading, error, callGraph } = useAuth();

const profile = ref<GraphProfile | null>(null);

const isClaimsError = computed(() =>
  error.value?.toLowerCase().includes("claims") ?? false
);

onMounted(async () => {
  const data = await callGraph<GraphProfile>("/me");
  if (data) {
    profile.value = data;
  }
});
</script>

<style scoped>
.back-link {
  display: inline-block;
  margin-bottom: 12px;
  color: #0078d4;
  text-decoration: none;
  font-size: 13px;
}
.back-link:hover {
  text-decoration: underline;
}

.card {
  background: #fff;
  border-radius: 8px;
  padding: 20px;
  box-shadow: 0 1px 4px rgba(0, 0, 0, 0.08);
}

.card h2 {
  font-size: 16px;
  margin-bottom: 12px;
}

.profile-table {
  width: 100%;
  border-collapse: collapse;
}

.profile-table td {
  padding: 6px 8px;
  font-size: 13px;
  border-bottom: 1px solid #edebe9;
}

.profile-table .label {
  font-weight: 600;
  width: 120px;
  color: #605e5c;
}

.raw {
  margin-top: 8px;
  font-size: 12px;
  background: #faf9f8;
  padding: 10px;
  border-radius: 4px;
  max-height: 200px;
  overflow: auto;
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

.error-banner .hint {
  margin-top: 6px;
  font-size: 12px;
  color: #605e5c;
}
</style>
