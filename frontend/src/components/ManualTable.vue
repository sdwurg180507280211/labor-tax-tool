<template>
  <div class="panel">
    <div class="manual-toolbar">
      <div>
        <h3>手动录入 / 从 Excel 粘贴</h3>
        <p>可以直接在表格里输入，也可以从 Excel 复制整块数据后粘贴到下方任意单元格。</p>
      </div>
      <div class="action-row">
        <button class="secondary" @click="addRow">新增一行</button>
        <button class="secondary" @click="clearRows">清空数据</button>
        <button class="primary" :disabled="loading" @click="submitRows">{{ loading ? '计算中...' : '开始计算' }}</button>
      </div>
    </div>

    <div class="table-wrap">
      <table class="input-table" @paste="handlePaste">
        <thead>
          <tr>
            <th v-for="col in columns" :key="col.key">{{ col.label }}<span v-if="col.required" class="required">*</span></th>
            <th>操作</th>
          </tr>
        </thead>
        <tbody>
          <tr v-for="(row, rowIndex) in rows" :key="row._id">
            <td v-for="col in columns" :key="col.key">
              <input v-model="row[col.key]" :type="col.type || 'text'" :placeholder="col.placeholder || ''" />
            </td>
            <td><button class="link-button" @click="removeRow(rowIndex)">删除</button></td>
          </tr>
        </tbody>
      </table>
    </div>
    <p class="hint">必填字段：年份、月份、讲者姓名、讲者ID、劳务费（实付金额）。</p>
  </div>
</template>

<script setup>
import { ref } from 'vue'

defineProps({ loading: Boolean })
const emit = defineEmits(['calculate'])

const columns = [
  { key: 'year', label: '年份', type: 'number', required: true, placeholder: '2026' },
  { key: 'month', label: '月份', type: 'number', required: true, placeholder: '6' },
  { key: 'department', label: '事业部' },
  { key: 'province', label: '省区' },
  { key: 'reimburser', label: '发起人/报销人' },
  { key: 'accountant', label: '会计' },
  { key: 'name', label: '讲者姓名', required: true },
  { key: 'id_no', label: '讲者ID', required: true },
  { key: 'after_tax_amount', label: '劳务费（实付金额）', type: 'number', required: true, placeholder: '3000' }
]

let idSeed = 1
function blankRow() {
  return { _id: idSeed++, year: new Date().getFullYear(), month: '', department: '', province: '', reimburser: '', accountant: '', name: '', id_no: '', after_tax_amount: '' }
}
const rows = ref([blankRow(), blankRow(), blankRow()])

function addRow() {
  rows.value.push(blankRow())
}

function removeRow(index) {
  rows.value.splice(index, 1)
  if (!rows.value.length) addRow()
}

function clearRows() {
  rows.value = [blankRow(), blankRow(), blankRow()]
}

function normalizeRows() {
  return rows.value
    .filter(row => columns.some(col => String(row[col.key] ?? '').trim() !== ''))
    .map(row => {
      const output = {}
      for (const col of columns) output[col.key] = row[col.key]
      return output
    })
}

function submitRows() {
  emit('calculate', normalizeRows())
}

function handlePaste(event) {
  const text = event.clipboardData?.getData('text')
  if (!text || !text.includes('\t')) return
  event.preventDefault()
  const lines = text.trim().split(/\r?\n/).filter(Boolean)
  const parsed = lines.map(line => line.split('\t'))
  rows.value = parsed.map(values => {
    const row = blankRow()
    columns.forEach((col, index) => {
      row[col.key] = values[index] ?? ''
    })
    return row
  })
}
</script>
