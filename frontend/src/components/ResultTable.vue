<template>
  <div class="table-wrap result-table-wrap">
    <table class="result-table">
      <thead>
        <tr>
          <th>序号</th>
          <th>年份</th>
          <th>月份</th>
          <th>讲者姓名</th>
          <th>讲者ID</th>
          <th>劳务费（实付金额）</th>
          <th>本月累计税后</th>
          <th>累计税前金额</th>
          <th>个税</th>
          <th>增值税</th>
          <th>附加税</th>
          <th>应开票金额</th>
          <th>应付款金额</th>
          <th>核对</th>
          <th v-if="debugMode">累计维度</th>
          <th v-if="debugMode">税后反推档位</th>
          <th v-if="debugMode">个税档位</th>
          <th v-if="debugMode">本次对应税前</th>
          <th v-if="debugMode">增值税规则</th>
          <th v-if="debugMode">跨档提示</th>
          <th v-if="debugMode">核对通过</th>
        </tr>
      </thead>
      <tbody>
        <tr v-for="row in rows" :key="row.index" :class="{ bad: Number(row.check_amount) !== 0 }">
          <td>{{ row.index }}</td>
          <td>{{ row.year }}</td>
          <td>{{ row.month }}</td>
          <td>{{ row.name }}</td>
          <td>{{ row.id_no }}</td>
          <td class="money">{{ row.after_tax_amount }}</td>
          <td class="money">{{ row.cumulative_after_tax_amount }}</td>
          <td class="money">{{ row.cumulative_pre_tax_without_vat }}</td>
          <td class="money">{{ row.individual_tax_amount }}</td>
          <td class="money">{{ row.vat_amount }}</td>
          <td class="money">{{ row.surcharge_amount }}</td>
          <td class="money">{{ row.invoice_amount }}</td>
          <td class="money">{{ row.payment_amount }}</td>
          <td class="money">{{ row.check_amount }}</td>
          <td v-if="debugMode">{{ row.debug_info?.cumulative_key }}</td>
          <td v-if="debugMode" class="debug-cell">{{ row.debug_info?.pre_tax_bracket }}</td>
          <td v-if="debugMode" class="debug-cell">{{ row.debug_info?.individual_tax_bracket }}</td>
          <td v-if="debugMode" class="money">{{ row.debug_info?.allocated_pre_tax_amount }}</td>
          <td v-if="debugMode" class="debug-cell">{{ row.debug_info?.vat_rule }}</td>
          <td v-if="debugMode">{{ row.debug_info?.threshold_note }}</td>
          <td v-if="debugMode">{{ row.debug_info?.check_ok }}</td>
        </tr>
      </tbody>
    </table>
  </div>
</template>

<script setup>
defineProps({
  rows: { type: Array, default: () => [] },
  debugMode: { type: Boolean, default: false }
})
</script>
