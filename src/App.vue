<template>
  <div>
    <el-button size="mini" type="warning" @click="exportExcel()">导出excel</el-button>
    <el-table stripe ref="report-table" border :data="tableData" id="export-table" style="width: 100%; color: black;"
      height="95vh">
      <el-table-column class="pink" type="index" width="50" label="序号" class-name="leave-alone" align="center">
      </el-table-column>
      <el-table-column class="pink" prop="shipper" label="发货人SHIPPER" align="center">
      </el-table-column>
      <el-table-column class="pink" prop="mblNo" align="center" label="主单号MBL">
      </el-table-column>
      <el-table-column class="pink" prop="containers" label="箱号Cont number" align="center">
      </el-table-column>
      <el-table-column class="pink" prop="hblNo" label="分单号HBL" align="center">
      </el-table-column>
      <el-table-column class="pink" prop="deliveryDate" label="送达日期 DELIVERY DATE" align="center">
      </el-table-column>
      <el-table-column class="pink" prop="goodsName" label="产品名称 NAME OF CARGO" align="center">
      </el-table-column>
      <el-table-column class="pink" prop="date" label="KIND OF CARGO" align="center">
      </el-table-column>
      <el-table-column class="pink" prop="date" label="合同号CONTRACT NO" align="center">
      </el-table-column>

      <el-table-column class="pink" prop="tradeType" label="贸易方式TRADE ITEM" align="center">
      </el-table-column>
      <el-table-column class="pink" prop="containerType" label="箱型量Cont quantity" align="center">
      </el-table-column>
      <!-- 1 -->
      <el-table-column class="pink" prop="date" label="瓶箱LCL" align="center">
      </el-table-column>
      <el-table-column class="pink" prop="ATA" label="ATA" align="center">
      </el-table-column>
      <el-table-column class="pink" prop="ATA" label="提货计划" align="center">
      </el-table-column>

      <el-table-column label="实报实销(应收)" align="center">
        <el-table-column prop="thc" label="码头操作费THC" align="center">
          <el-table-column prop="unusualFeeNote1" label="三级表头" align="center">
          </el-table-column>
          <el-table-column prop="unusualFeeNote1" label="三级表头1" align="center">
          </el-table-column>
          <!-- <template slot-scope="scope">{{scope.row.name ? scope.row.name : '--'}}</template> -->
        </el-table-column>
        <el-table-column prop="tlx" label="电放费TLX" align="center">
        </el-table-column>
        <el-table-column prop="mnf" label="开舱单费MNF" align="center">
        </el-table-column>
        <el-table-column prop="lss" label="低硫附加费LSS" align="center">
        </el-table-column>
        <el-table-column prop="cleanFee" label="一次洗箱费Cleaning fee" align="center">
        </el-table-column>
        <el-table-column prop="seal" label="锁费SEAL" align="center">
        </el-table-column>
        <el-table-column prop="piaoFee" label="票费Ticket Fee" align="center">
        </el-table-column>
        <el-table-column prop="emc" label="设备管理费EMC" align="center">
        </el-table-column>
        <el-table-column prop="cic" label="箱体不平衡费CIC" align="center">
        </el-table-column>
        <el-table-column prop="cfs" label="CFS" align="center">
        </el-table-column>
        <el-table-column prop="handingFee" label="Handling Fee" align="center">
        </el-table-column>
        <el-table-column prop="liftOn" label="吊上LIFT ON" align="center">
        </el-table-column>
      </el-table-column>
      <el-table-column prop="truckingFee" label="拖车费Trucking fee" align="center">
      </el-table-column>
      <el-table-column prop="operationFee" label="操作费" align="center">
      </el-table-column>
      <el-table-column label="实报实销" align="center">
        <el-table-column prop="saigonFee" label="胡志明港建费" align="center">
        </el-table-column>
        <el-table-column prop="delayFee" label="压车费" align="center">
        </el-table-column>
        <el-table-column prop="demFee" label="DEM" align="center">
        </el-table-column>
        <el-table-column prop="detFee" label="DET" align="center">
        </el-table-column>
        <el-table-column prop="stoFee" label="STO" align="center">
        </el-table-column>
        <el-table-column prop="repairingFee" label="修洗箱费REPAIRING FEE" align="center">
        </el-table-column>
        <el-table-column prop="unusualFee" label="异常费用" align="center">
        </el-table-column>
        <el-table-column prop="unusualFeeNote" label="异常费用原因" align="center">
        </el-table-column>
      </el-table-column>
      <el-table-column prop="customsClearance" label="清关费(USD)Customs Clearance" align="center">
      </el-table-column>
      <el-table-column prop="portSurcharge" label="海关小费Port Surcharge" align="center">
      </el-table-column>
      <el-table-column label="实报实销(VND)" align="center">
      </el-table-column>
      <el-table-column prop="exchange" label="汇率" align="center">
      </el-table-column>
      <el-table-column prop="incomeVnd" label="应收合计(VND)TOTAL" align="center">
      </el-table-column>
      <el-table-column label="应收合计(USD)TOTAL" align="center">
      </el-table-column>
      <el-table-column label="应付费用👉" align="center">
      </el-table-column>
      <el-table-column label="实报实销含税(应付)" align="center">
        <el-table-column prop="thcYf" label="码头操作费THC" align="center">
        </el-table-column>
        <el-table-column prop="tlxYf" label="电放费TLX" align="center">
        </el-table-column>
        <el-table-column prop="mnfYf" label="开舱单费MNF" align="center">
        </el-table-column>
        <el-table-column prop="lssYf" label="低硫附加费LSS " align="center">
        </el-table-column>
        <el-table-column prop="cleanFeeYf" label="一次洗箱费Cleaning fee" align="center">
        </el-table-column>
        <el-table-column prop="sealYf" label="锁费SEAL" align="center">
        </el-table-column>
        <el-table-column prop="piaoFeeYf" label="票费Ticket Fee" align="center">
        </el-table-column>
        <el-table-column prop="emcYf" label="设备管理费EMC" align="center">
        </el-table-column>
        <el-table-column prop="cicYf" label="箱体不平衡费CIC" align="center">
        </el-table-column>
        <el-table-column prop="cfsYf" label="CFS" align="center">
        </el-table-column>
        <el-table-column prop="handlingFeeYf" label="Handling Fee" align="center">
        </el-table-column>
        <el-table-column prop="liftOnYf" label="吊上LIFT ON" align="center">
        </el-table-column>
        <el-table-column prop="extraLoloYf" label="提柜附加费EXTRA LOLO" align="center">
        </el-table-column>
        <el-table-column prop="liftOffYf" label="吊下LIFT OFF" align="center">
        </el-table-column>
      </el-table-column>
      <el-table-column prop="truckingFeeYf" label="拖车Trucking fee" align="center">
      </el-table-column>
      <el-table-column prop="unLoadingFee" label="卸货" align="center">
      </el-table-column>
      <el-table-column prop="operationFeeYf" label="操作费" align="center">
      </el-table-column>
      <el-table-column label="实报实销含税(应付)" align="center">
        <el-table-column prop="saigonFeeYf" label="胡志明港建费" align="center">
        </el-table-column>
        <el-table-column prop="delayFeeYf" label="压车费" align="center">
        </el-table-column>
        <el-table-column prop="demFeeYf" label="DEM" align="center">
        </el-table-column>
        <el-table-column prop="detFeeYf" label="DET" align="center">
        </el-table-column>
        <el-table-column prop="stoFeeYf" label="STO" align="center">
        </el-table-column>
        <el-table-column prop="repairingFeeYf" label="修洗箱费REPAIRING FEE" align="center">
        </el-table-column>
        <el-table-column prop="unusualFeeYf" label="异常费用" align="center">
        </el-table-column>
        <el-table-column prop="unusualFeeNote" label="异常费用原因" align="center">
        </el-table-column>
      </el-table-column>
      <el-table-column prop="customsClearanceYf" label="清关费(USD)Customs Clearance" align="center">
      </el-table-column>
      <el-table-column prop="portSurchargeYf" label="海关小费Port Surcharge" align="center">
      </el-table-column>
      <el-table-column label="实报实销(VND)" align="center">
      </el-table-column>
      <el-table-column prop="exchange" label="汇率" align="center">
      </el-table-column>
      <el-table-column prop="outcomeVnd" label="应付合计(VND)Total" align="center">
      </el-table-column>
      <el-table-column prop="profit" label="业务毛利(VND)" align="center">
      </el-table-column>
    </el-table>
  </div>
  <!-- <img src="../packages/components/exportExcelStyle.js" alt=""> -->
</template>

<script>
import ExcelJS from 'exceljs';
import { exportExcelStyle } from '../packages/components/exportExcelStyle.js'
export default {
  data() {
    return {
      tableData: [],
    }
  },
  methods: {
    exportExcel() {
      // 设置表头样式
      const headerStyle = {
        alignment: {
          horizontal: 'center',
          vertical: 'center'
        },
        border: {
          top: { style: 'thin', color: 'black' },
          bottom: { style: 'thin', color: 'black' },
          left: { style: 'thin', color: 'black' },
          right: { style: 'thin', color: 'black' }
        }
      }
      // 设置普通单元格样式
      const cellStyle = {
        alignment: {
          horizontal: 'center',
          vertical: 'center'
        },
        border: {
          top: { style: 'thin', color: 'black' },
          bottom: { style: 'thin', color: 'black' },
          left: { style: 'thin', color: 'black' },
          right: { style: 'thin', color: 'black' }
        }
      }
      const tableDom = this.$refs['report-table'].$el;
      const name = 'report_example';
      // console.log('调用函数之前');

      exportExcelStyle(tableDom, headerStyle, cellStyle, name);
    }

  }
}
</script>

<style lang="scss" scoped></style>