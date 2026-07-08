<template>
  <div class="table-wrap result-table-wrap">
    <table class="result-table">
      <thead>
        <tr>
          <th>序号</th>
          <th>年份</th>
          <th>月份</th>
          <th>姓名</th>
          <th>身份证号码</th>
          <th>税后劳务金额</th>
          <th>本月累计税后</th>
          <th>累计税前金额</th>
          <th>个税</th>
          <th>增值税</th>
          <th>附加税</th>
          <th>应开票金额</th>
          <th>应付款金额</th>
          <th>核对</th>
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
        </tr>
      </tbody>
    </table>
  </div>
</template>

<script setup>
defineProps({ rows: { type: Array, default: () => [] } })
</script>
