<template>
  <div class="container">
    <el-card class="upload-card">
      <template #header>
        <div class="card-header">
          <h2>名单上传系统</h2>
        </div>
      </template>
      
      <el-upload
        class="upload-demo"
        drag
        action="#"
        :auto-upload="false"
        :on-change="handleFileChange"
        accept=".xlsx,.xls"
      >
        <el-icon class="el-icon--upload"><upload-filled /></el-icon>
        <div class="el-upload__text">
          将Excel文件拖到此处，或 <em>点击上传</em>
        </div>
        <template #tip>
          <div class="el-upload__tip">
            请上传Excel文件，要求包含序号和姓名列
          </div>
        </template>
      </el-upload>

      <div class="table-container">
        <div class="table-header">
          <h3>名单列表</h3>
          <el-button type="primary" @click="handleAdd">添加记录</el-button>
        </div>

        <el-table :data="tableData" style="width: 100%">
          <el-table-column prop="index" label="序号" width="180" />
          <el-table-column prop="name" label="姓名" width="180" />
          <el-table-column fixed="right" label="操作" width="180">
            <template #default="scope">
              <el-button link type="primary" @click="handleEdit(scope.row)">编辑</el-button>
              <el-button link type="danger" @click="handleDelete(scope.row)">删除</el-button>
            </template>
          </el-table-column>
        </el-table>
        
        <!-- 分页组件 -->
        <div class="pagination-container">
          <el-pagination
            v-model:current-page="currentPage"
            v-model:page-size="pageSize"
            :page-sizes="[10, 20, 50, 100]"
            layout="total, sizes, prev, pager, next"
            :total="total"
            @size-change="fetchData"
            @current-change="fetchData"
          />
        </div>
      </div>
    </el-card>

    <!-- 编辑对话框 -->
    <el-dialog v-model="dialogVisible" :title="dialogType === 'add' ? '添加记录' : '编辑记录'">
      <el-form :model="form" :rules="rules" ref="formRef">
        <el-form-item label="序号">
          <el-input v-model="form.index" />
        </el-form-item>
        <el-form-item label="姓名">
          <el-input v-model="form.name" />
        </el-form-item>
      </el-form>
      <template #footer>
        <span class="dialog-footer">
          <el-button @click="dialogVisible = false">取消</el-button>
          <el-button type="primary" @click="handleSave">确认</el-button>
        </span>
      </template>
    </el-dialog>
  </div>
</template>

<script setup>
import { ref, onMounted } from 'vue'
import { UploadFilled } from '@element-plus/icons-vue'
import { ElMessage } from 'element-plus'
import { read, utils } from 'xlsx'
import AV from 'leancloud-storage'

// 初始化LeanCloud
AV.init({
  appId: 'L0MSyA5ehj72PclaWpaWZaAG-gzGzoHsz',
  appKey: 'yueA3R2oKwScK3IizeJTFTpU',
  serverURL: 'https://l0msya5e.lc-cn-n1-shared.com'
})

const tableData = ref([])
const dialogVisible = ref(false)
const dialogType = ref('add')
const form = ref({
  index: '',
  name: '',
  objectId: null
})

// 分页相关数据
const currentPage = ref(1)
const pageSize = ref(10)
const total = ref(0)

// 表单验证规则
const rules = {
  index: [
    { required: true, message: '请输入序号', trigger: 'blur' },
    { type: 'number', message: '序号必须为数字', trigger: 'blur', transform: (value) => Number(value) }
  ],
  name: [{ required: true, message: '请输入姓名', trigger: 'blur' }]
}

// 文件上传处理
const handleFileChange = async (file) => {
  try {
    const data = await file.raw.arrayBuffer()
    const workbook = read(data)
    const worksheet = workbook.Sheets[workbook.SheetNames[0]]
    const jsonData = utils.sheet_to_json(worksheet)

    // 处理Excel数据
    const processedData = jsonData.map(row => ({
      index: Number(row['序号'] || row['序号'] || 0),
      name: row['姓名'] || row['姓名'] || ''
    }))

    // 保存到LeanCloud
    for (const item of processedData) {
      const Record = AV.Object.extend('Record')
      const record = new Record()
      await record.save(item)
    }

    // 刷新表格数据
    fetchData()
  } catch (error) {
    console.error('文件处理错误:', error)
    ElMessage.error('文件处理失败，请检查文件格式')
  }
}

// 获取数据
const fetchData = async () => {
  try {
    const query = new AV.Query('Record')
    // 设置分页
    query.skip((currentPage.value - 1) * pageSize.value)
    query.limit(pageSize.value)
    
    // 获取总数
    total.value = await query.count()
    
    const results = await query.find()
    tableData.value = results.map(result => ({
      objectId: result.id,
      index: result.get('index'),
      name: result.get('name')
    }))
  } catch (error) {
    console.error('获取数据失败:', error)
    ElMessage.error('获取数据失败')
  }
}

// 添加记录
const handleAdd = () => {
  dialogType.value = 'add'
  form.value = { index: '', name: '', objectId: null }
  dialogVisible.value = true
}

// 编辑记录
const handleEdit = (row) => {
  dialogType.value = 'edit'
  form.value = { ...row }
  dialogVisible.value = true
}

// 删除记录
const handleDelete = async (row) => {
  try {
    const record = AV.Object.createWithoutData('Record', row.objectId)
    await record.destroy()
    await fetchData()
    ElMessage.success('删除成功')
  } catch (error) {
    console.error('删除失败:', error)
    ElMessage.error('删除失败')
  }
}

// 保存记录
const formRef = ref(null)

const handleSave = async () => {
  if (!formRef.value) return
  
  try {
    await formRef.value.validate()
    
    if (dialogType.value === 'add') {
      const Record = AV.Object.extend('Record')
      const record = new Record()
      await record.save({
        index: Number(form.value.index),
        name: form.value.name
      })
    } else {
      const record = AV.Object.createWithoutData('Record', form.value.objectId)
      record.set('index', Number(form.value.index))
      record.set('name', form.value.name)
      await record.save()
    }
    
    dialogVisible.value = false
    await fetchData()
    ElMessage.success(dialogType.value === 'add' ? '添加成功' : '更新成功')
  } catch (error) {
    console.error('操作失败:', error)
    ElMessage.error(error.message || '操作失败')
  }
}

// 页面加载时获取数据
onMounted(() => {
  fetchData()
})
</script>

<style scoped>
.container {
  max-width: 1200px;
  margin: 20px auto;
  padding: 0 20px;
}

.upload-card {
  margin-bottom: 20px;
}

.card-header {
  display: flex;
  justify-content: space-between;
  align-items: center;
}

.table-container {
  margin-top: 20px;
}

.table-header {
  display: flex;
  justify-content: space-between;
  align-items: center;
  margin-bottom: 20px;
}

.upload-demo {
  margin: 20px 0;
}
.pagination-container {
  margin-top: 20px;
  display: flex;
  justify-content: flex-end;
}
</style>
