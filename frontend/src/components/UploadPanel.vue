<template>
  <div class="panel upload-panel">
    <div class="dropzone" @dragover.prevent @drop.prevent="onDrop">
      <input ref="fileInput" type="file" accept=".xlsx,.xlsm" hidden @change="onFileChange" />
      <div class="dropzone-icon">📄</div>
      <h3>上传新版简版0709劳务明细 Excel</h3>
      <p>支持 .xlsx / .xlsm。请包含：年份、月份、日期、姓名、讲者ID、税后劳务金额-单次。</p>
      <button class="secondary" type="button" @click="fileInput.click()">选择文件</button>
      <p v-if="selectedFile" class="file-name">已选择：{{ selectedFile.name }}</p>
    </div>
    <div class="action-row end">
      <button class="primary" :disabled="!selectedFile || loading" @click="emit('calculate', selectedFile)">
        {{ loading ? '计算中...' : '上传并计算' }}
      </button>
    </div>
  </div>
</template>

<script setup>
import { ref } from 'vue'

defineProps({ loading: Boolean })
const emit = defineEmits(['calculate'])
const fileInput = ref(null)
const selectedFile = ref(null)

function onFileChange(event) {
  selectedFile.value = event.target.files?.[0] || null
}

function onDrop(event) {
  selectedFile.value = event.dataTransfer.files?.[0] || null
}
</script>
