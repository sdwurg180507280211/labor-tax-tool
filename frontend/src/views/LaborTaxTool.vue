<template>
  <main class="page">
    <section class="hero">
      <div>
        <p class="eyebrow">无登录 · 不入库 · 上传即算 · 导出即走</p>
        <h1>劳务费税费换算工具</h1>
        <p class="desc">上传后台导出的劳务明细，系统按“年份 + 月份 + 讲者ID”累计，自动计算税前金额、个税、增值税、附加税、应开票金额和应付款金额。</p>
      </div>
      <div class="template-actions">
        <button class="secondary" @click="handleDownloadTemplate">下载导入模板</button>
        <button class="secondary" @click="handleDownloadLogicTestTemplate">下载逻辑测试模板</button>
        <button class="secondary danger-soft" @click="handleDownloadErrorTestTemplate">下载异常测试模板</button>
      </div>
    </section>

    <section class="card tool-options">
      <label class="test-mode-toggle">
        <input v-model="testMode" type="checkbox" />
        <span>
          <strong>测试模式</strong>
          <em>开启后，结果预览会显示累计维度、税后反推档位、个税档位、增值税规则和核对状态，方便排查公式逻辑。</em>
        </span>
      </label>
    </section>

    <section class="card">
      <div class="tabs">
        <button :class="['tab', activeTab === 'upload' && 'active']" @click="activeTab = 'upload'">上传 Excel</button>
        <button :class="['tab', activeTab === 'manual' && 'active']" @click="activeTab = 'manual'">手动录入 / 粘贴</button>
      </div>

      <UploadPanel v-if="activeTab === 'upload'" :loading="loading" @calculate="handleUploadCalculate" />
      <ManualTable v-else :loading="loading" @calculate="handleManualCalculate" />
    </section>

    <section v-if="error" class="alert error">{{ error }}</section>

    <section v-if="summary" class="summary-grid">
      <div class="summary-card"><span>数据行数</span><strong>{{ summary.row_count }}</strong></div>
      <div class="summary-card"><span>税后金额合计</span><strong>{{ summary.total_after_tax_amount }}</strong></div>
      <div class="summary-card"><span>应开票金额合计</span><strong>{{ summary.total_invoice_amount }}</strong></div>
      <div class="summary-card"><span>应付款金额合计</span><strong>{{ summary.total_payment_amount }}</strong></div>
      <div class="summary-card"><span>个税合计</span><strong>{{ summary.total_individual_tax_amount }}</strong></div>
    </section>

    <section v-if="resultRows.length" class="card result-card">
      <div class="result-actions">
        <div>
          <h2>计算结果预览</h2>
          <p>页面仅展示核心字段；导出 Excel 会包含“原表格式台账”“清晰版台账”“公式版台账”“测试核对报告”四个 Sheet。</p>
          <p v-if="testMode" class="debug-hint">测试模式已开启：当前表格额外显示每行计算档位和公式排查信息。</p>
        </div>
        <div class="action-row">
          <button class="secondary" @click="clearResult">清空结果</button>
          <button class="primary" :disabled="loading" @click="handleExport">导出完整台账 Excel</button>
        </div>
      </div>
      <ResultTable :rows="resultRows" :debug-mode="testMode" />
    </section>
  </main>
</template>

<script setup>
import { ref } from 'vue'
import UploadPanel from '../components/UploadPanel.vue'
import ManualTable from '../components/ManualTable.vue'
import ResultTable from '../components/ResultTable.vue'
import {
  calculateManual,
  calculateUpload,
  downloadErrorTestTemplate,
  downloadLogicTestTemplate,
  downloadTemplate,
  exportLedger
} from '../api/laborTaxApi'

const activeTab = ref('upload')
const loading = ref(false)
const error = ref('')
const resultRows = ref([])
const inputRows = ref([])
const summary = ref(null)
const testMode = ref(false)

async function runTask(task) {
  loading.value = true
  error.value = ''
  try {
    await task()
  } catch (err) {
    error.value = err.message || '操作失败'
  } finally {
    loading.value = false
  }
}

function applyResult(payload) {
  resultRows.value = payload.rows || []
  inputRows.value = payload.input_rows || []
  summary.value = payload.summary || null
}

function handleDownloadTemplate() {
  runTask(async () => downloadTemplate())
}

function handleDownloadLogicTestTemplate() {
  runTask(async () => downloadLogicTestTemplate())
}

function handleDownloadErrorTestTemplate() {
  runTask(async () => downloadErrorTestTemplate())
}

function handleUploadCalculate(file) {
  runTask(async () => {
    const payload = await calculateUpload(file)
    applyResult(payload)
  })
}

function handleManualCalculate(rows) {
  runTask(async () => {
    const payload = await calculateManual(rows)
    applyResult(payload)
  })
}

function handleExport() {
  runTask(async () => exportLedger(inputRows.value))
}

function clearResult() {
  resultRows.value = []
  inputRows.value = []
  summary.value = null
  error.value = ''
}
</script>
