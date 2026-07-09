<template>
  <div class="table-wrap result-table-wrap">
    <table class="result-table">
      <thead>
        <tr>
          <th>序号</th>
          <th>年份</th>
          <th>月份</th>
          <th>日期</th>
          <th>姓名</th>
          <th>讲者ID</th>
          <th>税后劳务金额-单次</th>
          <th>本月累计税后</th>
          <th>累计税前金额</th>
          <th>单日累计税前</th>
          <th>协议签订金额</th>
          <th>增值税</th>
          <th>附加税</th>
          <th>个税</th>
          <th v-if="debugMode">月累计维度</th>
          <th v-if="debugMode">日累计维度</th>
          <th v-if="debugMode">税后反推档位</th>
          <th v-if="debugMode">个税档位</th>
          <th v-if="debugMode">增值税规则</th>
        </tr>
      </thead>
      <tbody>
        <tr v-for="row in rows" :key="row.index">
          <td>{{ row.index }}</td>
          <td>{{ row.year }}</td>
          <td>{{ row.month }}</td>
          <td>{{ row.day }}</td>
          <td>{{ row.name }}</td>
          <td>{{ row.id_no }}</td>
          <td class="money">{{ row.after_tax_amount }}</td>
          <td class="money">{{ row.cumulative_after_tax_amount }}</td>
          <td class="money">{{ row.cumulative_pre_tax_without_vat }}</td>
          <td class="money">{{ row.daily_cumulative_pre_tax_without_vat }}</td>
          <td class="money">{{ row.contract_amount }}</td>
          <td class="money">{{ row.vat_amount }}</td>
          <td class="money">{{ row.surcharge_amount }}</td>
          <td class="money">{{ row.individual_tax_amount }}</td>
          <td v-if="debugMode">{{ row.debug_info?.monthly_cumulative_key }}</td>
          <td v-if="debugMode">{{ row.debug_info?.daily_cumulative_key }}</td>
          <td v-if="debugMode" class="debug-cell">{{ row.debug_info?.pre_tax_bracket }}</td>
          <td v-if="debugMode" class="debug-cell">{{ row.debug_info?.individual_tax_bracket }}</td>
          <td v-if="debugMode" class="debug-cell">{{ row.debug_info?.vat_rule }}</td>
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
