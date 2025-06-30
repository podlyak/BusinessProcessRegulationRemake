import groovy.util.logging.Slf4j
import org.apache.poi.ooxml.POIXMLRelation
import org.apache.poi.openxml4j.opc.PackagePart
import org.apache.poi.openxml4j.opc.PackageRelationshipTypes
import org.apache.poi.util.IOUtils
import org.apache.poi.util.Units
import org.apache.poi.xwpf.usermodel.BodyElementType
import org.apache.poi.xwpf.usermodel.IBody
import org.apache.poi.xwpf.usermodel.IBodyElement
import org.apache.poi.xwpf.usermodel.ParagraphAlignment
import org.apache.poi.xwpf.usermodel.UnderlinePatterns
import org.apache.poi.xwpf.usermodel.XWPFAbstractNum
import org.apache.poi.xwpf.usermodel.XWPFDocument
import org.apache.poi.xwpf.usermodel.XWPFFactory
import org.apache.poi.xwpf.usermodel.XWPFHyperlinkRun
import org.apache.poi.xwpf.usermodel.XWPFNum
import org.apache.poi.xwpf.usermodel.XWPFNumbering
import org.apache.poi.xwpf.usermodel.XWPFParagraph
import org.apache.poi.xwpf.usermodel.XWPFPictureData
import org.apache.poi.xwpf.usermodel.XWPFRun
import org.apache.poi.xwpf.usermodel.XWPFTable
import org.apache.poi.xwpf.usermodel.XWPFTableCell
import org.apache.poi.xwpf.usermodel.XWPFTableRow
import org.apache.xmlbeans.XmlCursor
import org.apache.xmlbeans.XmlToken
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTAbstractNum
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTBookmark
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTDecimalNumber
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTFonts
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTHyperlink
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTLvl
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTNumLvl
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTPPr
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTPageSz
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTR
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTRPr
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTSectPr
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTShd
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTSimpleField
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STJc
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STNumberFormat
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STPageOrientation
import ru.nextconsulting.bpm.dto.FullModelDefinition
import ru.nextconsulting.bpm.dto.NodeId
import ru.nextconsulting.bpm.dto.SimpleMultipartFile
import ru.nextconsulting.bpm.repository.business.AttributeValue
import ru.nextconsulting.bpm.repository.model.layout.GridCell
import ru.nextconsulting.bpm.repository.model.layout.GridHeader
import ru.nextconsulting.bpm.repository.structure.FileNodeDTO
import ru.nextconsulting.bpm.repository.structure.ModelNode
import ru.nextconsulting.bpm.repository.structure.Node
import ru.nextconsulting.bpm.repository.structure.ObjectDefinitionNode
import ru.nextconsulting.bpm.repository.structure.ScriptParameter
import ru.nextconsulting.bpm.repository.structure.SilaScriptParamType
import ru.nextconsulting.bpm.script.repository.TreeRepository
import ru.nextconsulting.bpm.script.tree.elements.Edge
import ru.nextconsulting.bpm.script.tree.elements.ObjectElement
import ru.nextconsulting.bpm.script.tree.node.Model
import ru.nextconsulting.bpm.script.tree.node.ObjectDefinition
import ru.nextconsulting.bpm.script.tree.node.TreeNode
import ru.nextconsulting.bpm.script.utils.ModelUtils
import ru.nextconsulting.bpm.scriptengine.context.ContextParameters
import ru.nextconsulting.bpm.scriptengine.context.CustomScriptContext
import ru.nextconsulting.bpm.scriptengine.customapi.ImageApi
import ru.nextconsulting.bpm.scriptengine.exception.SilaScriptException
import ru.nextconsulting.bpm.scriptengine.script.GroovyScript
import ru.nextconsulting.bpm.scriptengine.serverapi.FileApi
import ru.nextconsulting.bpm.scriptengine.serverapi.ModelApi
import ru.nextconsulting.bpm.scriptengine.util.ParamUtils
import ru.nextconsulting.bpm.scriptengine.util.SilaScriptParameter
import ru.nextconsulting.bpm.scriptengine.util.SilaScriptParameters
import ru.nextconsulting.bpm.utils.JsonConverter

import javax.imageio.ImageIO
import java.awt.image.BufferedImage
import java.sql.Timestamp
import java.text.SimpleDateFormat
import java.time.LocalDate
import java.util.regex.Matcher
import java.util.regex.Pattern
import java.util.zip.ZipEntry
import java.util.zip.ZipOutputStream

@SuppressWarnings('unused')
void execute() {
    new BusinessProcessRegulationRemakeScript(context: context).execute()
}

@SilaScriptParameters([
        @SilaScriptParameter(
                name = DETAIL_LEVEL_PARAM_NAME,
                type = SilaScriptParamType.SELECT_STRING,
                selectStringValues = ['3 уровень', '4 уровень'],
                defaultValue = '3 уровень'
        ),
        @SilaScriptParameter(
                name = DOC_VERSION_PARAM_NAME,
                type = SilaScriptParamType.STRING,
                required = false,
                defaultValue = ''
        ),
        @SilaScriptParameter(
                name = DOC_DATE_PARAM_NAME,
                type = SilaScriptParamType.DATE,
                required = false,
                defaultValue = ''
        ),
        @SilaScriptParameter(
                name = IMAGE_TYPE_PARAM_NAME,
                type = SilaScriptParamType.SELECT_STRING,
                selectStringValues = ['PNG', 'SVG'],
                defaultValue = 'PNG'
        ),
])
@Slf4j
class BusinessProcessRegulationRemakeScript implements GroovyScript {
    static void main(String[] args) {
        ContextParameters parameters = ContextParameters.builder()
                .login('superadmin')
                .password('WM_Sila_123')
                .apiBaseUrl('http://localhost:8080/')
                .imageServiceUrl('http://localhost:8084/')
                .silaUrl('http://localhost:8080/')
                .build()
        CustomScriptContext context = CustomScriptContext.create(parameters)

        ScriptParameter modelParam = ScriptParameter.builder()
                .paramType(SilaScriptParamType.NODE)
                .name('modelId')
                .value(JsonConverter.writeValueAsJson(NodeId.builder()
                        .repositoryId('51b21ba6-c89c-49e2-911e-9c88b609b728')
                        .id('9df27520-b000-11e6-05b7-db7cafd96ef7')
                        .build())
                )
                .build()
        ScriptParameter elementsIdsParam = ScriptParameter.builder()
                .paramType(SilaScriptParamType.STRING_LIST)
                .name('elementsIdsList')
                .value('["72c58d3e-b261-11e6-05b7-db7cafd96ef7"]')
                .build()

        context.getParameters().add(modelParam)
        context.getParameters().add(elementsIdsParam)

        BusinessProcessRegulationRemakeScript script = new BusinessProcessRegulationRemakeScript(context: context)
        script.execute()
    }

    static final String DETAIL_LEVEL_PARAM_NAME = 'Глубина детализации регламента'
    static final String DOC_VERSION_PARAM_NAME = 'Номер версии регламента'
    static final String DOC_DATE_PARAM_NAME = 'Дата утверждения регламента'
    static final String IMAGE_TYPE_PARAM_NAME = 'Формат получаемых изображений'

    //------------------------------------------------------------------------------------------------------------------
    // константы id элементов
    //------------------------------------------------------------------------------------------------------------------
    private static final String ABBREVIATIONS_MODEL_ID = '0c25ad70-2733-11e6-05b7-db7cafd96ef7'
    private static final String ABBREVIATIONS_ROOT_OBJECT_ID = '0f7107e4-2733-11e6-05b7-db7cafd96ef7'
    private static final String FILE_REPOSITORY_ID = 'file-folder-root-id'
    private static final String FIRST_LEVEL_MODEL_ID = '1a8132f0-a43b-11e7-05b7-db7cafd96ef7'

    //------------------------------------------------------------------------------------------------------------------
    // константы для работы с файлами
    //------------------------------------------------------------------------------------------------------------------
    private static final String DOCX_RESULT_FILE_NAME_FIRST_PART = 'Регламент БП'
    private static final String ZIP_RESULT_FILE_NAME_FIRST_PART = 'Регламенты БП'
    private static final String DOCX_FORMAT = 'docx'
    private static final String ZIP_FORMAT = 'zip'
    private static final String BUSINESS_PROCESS_REGULATION_TEMPLATE_NAME = 'business_process_regulation_template_v11.docx'
    private static final String TEMPLATE_FOLDER_NAME = 'Общие'

    //------------------------------------------------------------------------------------------------------------------
    // константы для отладки при разработке
    //------------------------------------------------------------------------------------------------------------------
    private static final boolean DEBUG = true
    private static final String TEMPLATE_LOCAL_PATH = 'C:\\Users\\vikto\\IdeaProjects\\BusinessProcessRegulationRemake\\examples'

    //------------------------------------------------------------------------------------------------------------------
    // константы шаблона для простого текста
    //------------------------------------------------------------------------------------------------------------------
    private static final String PROCESS_NAME_UPPER_CASE_TEMPLATE_KEY = 'НАЗВАНИЕ ПРОЦЕССА'
    private static final String PROCESS_CODE_TEMPLATE_KEY = 'Код процесса'
    private static final String DOC_VERSION_WITH_DATE_TEMPLATE_KEY = 'XX от XX.XX.XXXX'
    private static final String DOC_YEAR_TEMPLATE_KEY = 'Год'
    private static final String DOC_VERSION_TEMPLATE_KEY = 'XX'
    private static final String DOC_DATE_TEMPLATE_KEY = 'XX.XX.XXXX'
    private static final String PROCESS_NAME_TEMPLATE_KEY = 'Наименование процесса'
    private static final String FIRST_LEVEL_PROCESS_CODE_TEMPLATE_KEY = 'Код ПВУ'
    private static final String FIRST_LEVEL_PROCESS_NAME_TEMPLATE_KEY = 'ПВУ'
    private static final String PROCESS_REQUIREMENTS_TEMPLATE_KEY = 'Требования к процессу'

    //------------------------------------------------------------------------------------------------------------------
    // константы шаблона для списков
    //------------------------------------------------------------------------------------------------------------------
    private static final String PROCESS_OWNER_POSITION_TEMPLATE_KEY = 'Должность владельца процесса'
    private static final String PROCESS_OWNER_TEMPLATE_KEY = 'Владелец процесса'
    private static final String REQUISITES_NORMATIVE_DOCUMENT_TEMPLATE_KEY = 'Реквизиты нормативного документа'
    private static final String PROCESS_GOAL_TEMPLATE_KEY = 'Цель процесса'

    //------------------------------------------------------------------------------------------------------------------
    // константы заголовков таблиц
    //------------------------------------------------------------------------------------------------------------------
    private static final List<String> ABBREVIATION_TABLE_HEADERS = [
            'Сокращение',
            'Расшифровка',
    ]
    private static final List<String> BUSINESS_PROCESS_HIERARCHY_TABLE_HEADERS = [
            'Уровень',
            'Наименование процесса',
            'Раздел',
    ]
    private static final List<String> EXTERNAL_BUSINESS_PROCESS_TABLE_HEADERS = [
            'Смежный процесс',
            'Вход из смежного процесса',
            'Выход в смежный процесс',
    ]
    private static final List<String> PROCESS_BUSINESS_ROLE_TABLE_HEADERS = [
            'Роль',
            'Должности',
    ]
    private static final List<String> DOCUMENT_COLLECTION_TABLE_HEADERS = [
            'Набор документов',
            'Состав набора документов',
    ]
    private static final List<String> FUNCTIONS_TABLE_HEADERS = [
            "<${FUNCTION_NUMBER_TEMPLATE_KEY}>",
            "<${FUNCTION_CODE_TEMPLATE_KEY}> <${FUNCTION_NAME_TEMPLATE_KEY}>",
    ]
    private static final List<String> FUNCTIONS_TABLE_DOCUMENT_HORIZONTAL_HEADERS = [
            'Входы',
            'Выходы',
    ]
    private static final List<String> RESPONSIBILITY_MATRIX_HEADERS = [
            'Процедура / Роль',
            "<${BUSINESS_ROLE_NAME_TEMPLATE_KEY}>",
    ]

    //------------------------------------------------------------------------------------------------------------------
    // константы частей текста искомых параграфов
    //------------------------------------------------------------------------------------------------------------------
    private static final String PROCESS_BUSINESS_ROLES_PARAGRAPH_TEXT_TO_FIND = 'Перечень ролей Процесса'
    private static final String PROCESS_PARTICIPANTS_PARAGRAPH_TEXT_TO_FIND = 'Участники процесса указаны в таблицах'
    private static final String SCENARIO_PARAGRAPH_TEXT_TO_FIND = "Сценарий <${SCENARIO_CODE_TEMPLATE_KEY}> <${SCENARIO_NAME_TEMPLATE_KEY}>"
    private static final String DOCUMENT_COLLECTION_PARAGRAPH_TEXT_TO_FIND = 'Состав наборов документов'
    private static final String REQUIREMENTS_PARAGRAPH_TEXT_TO_FIND = 'Требования к'
    private static final String PICTURE_PARAGRAPH_TEXT_TO_FIND = 'На рисунке'
    private static final String FUNCTIONS_PARAGRAPH_TEXT_TO_FIND = 'Порядок взаимодействия участников'
    private static final String RESPONSIBILITY_MATRIX_PARAGRAPH_TEXT_TO_FIND = 'Матрица ответственности сценария приведена в таблице'
    private static final String SCENARIO_FUNCTIONS_TABLE_PARAGRAPH_TEXT_TO_FIND = 'Порядок взаимодействия участников сценария описан в таблице'
    private static final String PROCEDURE_PARAGRAPH_TEXT_TO_FIND = "Процедура <${PROCEDURE_CODE_TEMPLATE_KEY}> <${PROCEDURE_NAME_TEMPLATE_KEY}>"

    //------------------------------------------------------------------------------------------------------------------
    // константы шаблона для таблиц
    //------------------------------------------------------------------------------------------------------------------
    private static final String ABBREVIATION_TEMPLATE_KEY = 'Сокращение'
    private static final String ABBREVIATION_VALUE_TEMPLATE_KEY = 'Значение сокращения'
    private static final String BUSINESS_PROCESS_LEVEL_TEMPLATE_KEY = 'Номер уровня БП'
    private static final String BUSINESS_PROCESS_CODE_TEMPLATE_KEY = 'Код БП'
    private static final String BUSINESS_PROCESS_NAME_TEMPLATE_KEY = 'Наименование БП'
    private static final String EXTERNAL_BUSINESS_PROCESS_CODE_TEMPLATE_KEY = 'Код смежного БП'
    private static final String EXTERNAL_BUSINESS_PROCESS_NAME_TEMPLATE_KEY = 'Смежный БП'
    private static final String EXTERNAL_BUSINESS_PROCESS_INPUT_TEMPLATE_KEY = 'Вход из смежного БП'
    private static final String EXTERNAL_BUSINESS_PROCESS_OUTPUT_TEMPLATE_KEY = 'Выход в смежный БП'
    private static final String PROCESS_BUSINESS_ROLE_TEMPLATE_KEY = 'Роль процесса'
    private static final String PROCESS_BUSINESS_ROLE_POSITION_TEMPLATE_KEY = 'Должность для роли'
    private static final String PROCESS_BUSINESS_ROLE_POSITION_ORGANIZATIONAL_UNIT_TEMPLATE_KEY = 'ОЕ для должности'
    private static final String PROCESS_DOCUMENT_COLLECTION_TEMPLATE_KEY = 'Набор документов'
    private static final String PROCESS_DOCUMENT_COLLECTION_CONTAINED_DOCUMENT_TEMPLATE_KEY = 'Документы набора'

    //------------------------------------------------------------------------------------------------------------------
    // константы шаблона для раздела моделей
    //------------------------------------------------------------------------------------------------------------------
    private static final String PROCESS_MODEL_TEMPLATE_KEY = 'Модель процесса'
    private static final String SCENARIO_CODE_TEMPLATE_KEY = 'Код сценария'
    private static final String SCENARIO_NAME_TEMPLATE_KEY = 'Сценарий'
    private static final String SCENARIO_REQUIREMENTS_TEMPLATE_KEY = 'Требования к сценарию'
    private static final String SCENARIO_MODEL_TEMPLATE_KEY = 'Модель сценария'
    private static final String FUNCTION_NUMBER_TEMPLATE_KEY = '№ пп'
    private static final String FUNCTION_CODE_TEMPLATE_KEY = 'Код функции'
    private static final String FUNCTION_NAME_TEMPLATE_KEY = 'Функция'
    private static final String PERFORMER_TEMPLATE_KEY = 'Исполнитель'
    private static final String INPUT_DOCUMENT_TEMPLATE_KEY = 'Входящий документ'
    private static final String INPUT_EVENT_TEMPLATE_KEY = 'Входящее событие'
    private static final String INFORMATION_SYSTEM_TEMPLATE_KEY = 'Информационная система'
    private static final String FUNCTION_REQUIREMENTS_TEMPLATE_KEY = 'Требования к функции'
    private static final String OUTPUT_DOCUMENT_TEMPLATE_KEY = 'Исходящий документ'
    private static final String OUTPUT_EVENT_TEMPLATE_KEY = 'Исходящее событие'
    private static final String DURATION_TEMPLATE_KEY = 'Длительность'
    private static final String CHILD_FUNCTION_TEMPLATE_KEY = 'Условие'
    private static final String BUSINESS_ROLE_NAME_TEMPLATE_KEY = 'Роль'
    private static final String PROCEDURE_CODE_TEMPLATE_KEY = 'Код процедуры'
    private static final String PROCEDURE_NAME_TEMPLATE_KEY = 'Процедура'
    private static final String PROCEDURE_REQUIREMENTS_TEMPLATE_KEY = 'Требования к процедуре'
    private static final String PROCEDURE_MODEL_TEMPLATE_KEY = 'Модель процедуры'

    //------------------------------------------------------------------------------------------------------------------
    // константы шаблона для номеров рисунков и таблиц
    //------------------------------------------------------------------------------------------------------------------
    private static final String SCENARIO_PICTURE_NUMBER_TEMPLATE_KEY = 'Номер рисунка сценария'
    private static final String SCENARIO_FUNCTIONS_TABLE_NUMBER_TEMPLATE_KEY = 'Номер таблицы порядок взаимодействия участников сценария'
    private static final String SCENARIO_RESPONSIBILITY_MATRIX_TABLE_NUMBER_TEMPLATE_KEY = 'Номер таблицы матрица ответственности сценария'
    private static final String PROCEDURE_PICTURE_NUMBER_TEMPLATE_KEY = 'Номер рисунка процедуры'
    private static final String PROCEDURE_FUNCTIONS_TABLE_NUMBER_TEMPLATE_KEY = 'Номер таблицы порядок взаимодействия участников процедуры'

    //------------------------------------------------------------------------------------------------------------------
    // константы id элементов
    //------------------------------------------------------------------------------------------------------------------
    private static final String FILE_NODE_TYPE_ID = 'FILE_FOLDER'

    private static final String EPC_MODEL_TYPE_ID = 'MT_EEPC'
    private static final String FUNCTION_ALLOCATION_MODEL_TYPE_ID = 'MT_FUNC_ALLOC_DGM'
    private static final String IEF_DATA_MODEL_TYPE_ID = 'MT_IEF_DATA_MDL'
    private static final String INFORMATION_CARRIER_MODEL_TYPE_ID = 'MT_INFO_CARR_DGM'
    private static final String ORGANIZATION_STRUCTURE_MODEL_TYPE_ID = 'MT_ORG_CHRT'
    private static final String PROCESS_SELECTION_MODEL_TYPE_ID = 'MT_PRCS_SLCT_DIA'

    private static final List<String> DOCUMENT_COLLECTION_MODEL_TYPE_IDS = [
            IEF_DATA_MODEL_TYPE_ID,
            INFORMATION_CARRIER_MODEL_TYPE_ID,
    ]

    private static final String APPLICATION_SYSTEM_TYPE_OBJECT_TYPE_ID = 'OT_APPL_SYS_TYPE'
    private static final String BUSINESS_ROLE_OBJECT_TYPE_ID = 'OT_PERS_TYPE'
    private static final String CLUSTER_DATA_MODEL_OBJECT_TYPE_ID = 'OT_CLST'
    private static final String EVENT_OBJECT_TYPE_ID = 'OT_EVT'
    private static final String FLOW_OBJECT_TYPE_ID = 'OT_TECH_TRM'
    private static final String FUNCTION_OBJECT_TYPE_ID = 'OT_FUNC'
    private static final String GOAL_OBJECT_TYPE_ID = 'OT_OBJECTIVE'
    private static final String GROUP_OBJECT_TYPE_ID = 'OT_GRP'
    private static final String INFORMATION_CARRIER_OBJECT_TYPE_ID = 'OT_INFO_CARR'
    private static final String ORGANIZATIONAL_UNIT_OBJECT_TYPE_ID = 'OT_ORG_UNIT'
    private static final String RULE_OBJECT_TYPE_ID = 'OT_RULE'

    private static final List<String> DOCUMENT_OBJECT_TYPE_IDS = [
            CLUSTER_DATA_MODEL_OBJECT_TYPE_ID,
            INFORMATION_CARRIER_OBJECT_TYPE_ID,
    ]

    private static final List<String> ABBREVIATION_EDGE_TYPE_IDS = [
            'CT_HAS_REL_WITH',
            'CT_IS_IN_RELSHP_TO',
            'CT_IS_IN_RELSHP_TO_1',
            'CT_REFS_TO_2',
    ]
    private static final List<String> CLUSTER_GROUP_W_CLUSTER_DATA_MODEL_EDGE_TYPE_IDS = [
            'CT_CONS_OF_1',
            'CT_CONS_OF_2',
    ]
    private static final List<String> DOCUMENT_COLLECTION_W_DOCUMENT_EDGE_TYPE_IDS = [
            'CT_CAN_SUBS_2',
            'CT_SUBS_1',
            'CT_SUBS_3',
            'CT_SUBS_5',
    ]
    private static final List<String> DOCUMENT_W_EPC_FUNCTION_EDGE_TYPE_IDS = [
            'CT_IS_INP_FOR',
            'CT_PROV_INP_FOR',
    ]
    private static final String DOCUMENT_W_STATUS_EDGE_TYPE_ID = 'CT_HAS_STATE'
    private static final List<String> EPC_FUNCTION_W_DOCUMENT_EDGE_TYPE_IDS = [
            'CT_CRT_OUT_TO',
            'CT_HAS_OUT',
    ]
    private static final List<String> EPC_FUNCTION_W_EVENT_EDGE_TYPE_IDS = [
            'CT_CRT_1',
            'CT_CRT_3',
    ]
    private static final List<String> EPC_FUNCTION_W_OPERATOR_EDGE_TYPE_IDS = [
            'CT_LEADS_TO_1',
            'CT_LEADS_TO_2',
    ]
    private static final String EVENT_W_EPC_FUNCTION_EDGE_TYPE_ID = 'CT_ACTIV_1'
    private static final String EVENT_W_OPERATOR_EDGE_TYPE_ID = 'CT_IS_EVAL_BY_1'
    private static final List<String> INFORMATION_SYSTEM_W_EPC_FUNCTION_EDGE_TYPE_IDS = [
            'CT_CAN_SUPP_1',
            'CT_SUPP_1',
            'CT_SUPP_2',
            'CT_SUPP_3',
    ]
    private static final String INPUT_FLOW_W_SUBPROCESS_EDGE_TYPE_ID = 'CT_IS_INP_FOR'
    private static final String LEADERSHIP_POSITION_W_OWNER_EDGE_TYPE_ID = 'CT_IS_DISC_SUPER'
    private static final String OPERATOR_W_EPC_FUNCTION_EDGE_TYPE_ID = 'CT_ACTIV_1'
    private static final List<String> OPERATOR_W_EVENT_EDGE_TYPE_IDS = [
            'CT_LEADS_TO_1',
            'CT_LEADS_TO_2',
    ]
    private static final List<String> OPERATOR_W_OPERATOR_EDGE_TYPE_IDS = [
            'CT_BPEL_LINKS',
            'CT_LNK_1',
            'CT_LNK_2',
    ]
    private static final String ORGANIZATIONAL_UNIT_W_POSITION_EDGE_TYPE_ID = 'CT_IS_CRT_BY'
    private static final String OUTPUT_FLOW_W_CUSTOMER_EDGE_TYPE_ID = 'CT_IS_INP_FOR'
    private static final List<String> OWNER_W_SUBPROCESS_EDGE_TYPE_IDS = [
            'CT_EXEC_1',
            'CT_EXEC_2',
    ]
    private static final List<String> PERFORMER_W_EPC_FUNCTION_EDGE_TYPE_IDS = [
            'CT_AGREES',
            'CT_CONTR_TO_1',
            'CT_CONTR_TO_2',
            'CT_DECD_ON',
            'CT_EXEC_1',
            'CT_EXEC_2',
            'CT_MUST_BE_INFO_ABT_1',
    ]
    private static final String POSITION_W_BUSINESS_ROLE_EDGE_TYPE_ID = 'CT_EXEC_5'
    private static final String SUBPROCESS_W_OUTPUT_FLOW_EDGE_TYPE_ID = 'CT_HAS_OUT'
    private static final String SUPPLIER_W_INPUT_FLOW_EDGE_TYPE_ID = 'CT_HAS_OUT'

    private static final String AVERAGE_EXECUTION_TIME_ATTR_ID = 'AT_TIME_AVG_PRCS'
    private static final String DATA_ELEMENT_CODE_ATTR_ID = '46e148b0-b96d-11e3-05b7-db7cafd96ef7'
    private static final String DESCRIPTION_DEFINITION_ATTR_ID = 'AT_DESC'
    private static final String FULL_NAME_ATTR_ID = 'AT_NAME_FULL'

    private static final List<String> EXCLUDED_FUNCTION_SYMBOL_IDS = [
            '75f2e570-bdd3-11e5-05b7-db7cafd96ef7', // интерфейс смежного процесса
            'ST_PRCS_IF', // интерфейс процесса
            'fd841c20-cc37-11e6-05b7-db7cafd96ef7', // группировка интерфейсов
    ]
    private static final String EXTERNAL_PROCESS_SYMBOL_ID = '75d9e6f0-4d1a-11e3-58a3-928422d47a25'
    private static final String NORMATIVE_DOCUMENT_SYMBOL_ID = '7096d320-cf42-11e2-69e4-ac8112d1b401'
    private static final List<String> SCENARIO_SYMBOL_IDS = [
            '1647b400-c1a5-11e4-3864-ff0f8fe73e88', // сценарий SAP (типовой)
            '1bea43c0-c768-11e2-69e4-ac8112d1b401', // сценарий (типовой)
            '478e24e0-c1a5-11e4-3864-ff0f8fe73e88', // сценарий SAP
            'ST_SCENARIO', // сценарий
    ]
    private static final String STATUS_SYMBOL_ID = 'd6e8a7b0-7ce6-11e2-3463-e4115bf4fdb9'

    //------------------------------------------------------------------------------------------------------------------
    // основной код
    //------------------------------------------------------------------------------------------------------------------
    private static final int CM_1_OFFSET = 567 // 1 сантиметр (отступ в документе)

    private static Map<String, String> fullAbbreviations = new TreeMap<>()
    private static Pattern abbreviationsPattern = null
    private static Map<String, String> foundedAbbreviations = new TreeMap<>()

    CustomScriptContext context
    ImageApi imageApi
    ModelApi modelApi
    private TreeRepository treeRepository

    private static int detailLevel
    private static String docVersion
    private static String docDate
    private static ImageType imageType
    private static String currentYear = LocalDate.now().getYear().toString()

    enum ImageType {
        PNG,
        SVG,
    }

    class ModelImage {
        byte[] image
        ImageType type
        int width
        int height

        ModelImage(byte[] image, ImageType type, int width, int height) {
            this.image = image
            this.type = type
            this.width = width
            this.height = height
        }

        void addToRun(XWPFDocument document, XWPFRun run, int width, int height) {
            if (type == ImageType.PNG) {
                InputStream is = new ByteArrayInputStream(image)
                run.addPicture(is, XWPFDocument.PICTURE_TYPE_PNG, "image.png", Units.toEMU(width), Units.toEMU(height))
            }

            if (type == ImageType.SVG) {
                CTR ctr = run.getCTR()

                String blipIdPng = document.addPictureData(new ByteArrayInputStream(image), XWPFDocument.PICTURE_TYPE_PNG)
                int id = document.getNextPicNameNumber(XWPFDocument.PICTURE_TYPE_PNG)
                String blipId = ((XWPFDocumentSvg) document).addSVGPicture(new ByteArrayInputStream(image))
                String widthString = String.valueOf(Units.pixelToEMU(width))
                String heightString = String.valueOf(Units.pixelToEMU(height))
                String xmlSvgStructure = "<w:drawing\n" +
                        "xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" xmlns:rel=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">" +
                        "<wp:inline\n" +
                        "xmlns:wp=\"http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing\" distT=\"0\" distB=\"0\" distL=\"0\" distR=\"0\">" +
                        "<wp:extent cx=\"" + widthString + "\" cy=\"" + heightString + "\"/>" +
                        "<wp:effectExtent l=\"0\" t=\"0\" r=\"9525\" b=\"9525\"/>" +
                        "<wp:docPr id=\"1\" name=\"Graphic 1\"/>" +
                        "<wp:cNvGraphicFramePr>" +
                        "<a:graphicFrameLocks\n" +
                        "xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" noChangeAspect=\"1\"/>" +
                        "</wp:cNvGraphicFramePr>" +
                        "<a:graphic\n" +
                        "xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\">" +
                        "<a:graphicData uri=\"http://schemas.openxmlformats.org/drawingml/2006/picture\">" +
                        "<pic:pic\n" +
                        "xmlns:pic=\"http://schemas.openxmlformats.org/drawingml/2006/picture\">" +
                        "<pic:nvPicPr>" +
                        "<pic:cNvPr id=\"" + id + "\" name=\"test.svg\"/>" +
                        "<pic:cNvPicPr/>" +
                        "</pic:nvPicPr>" +
                        "<pic:blipFill>" +
                        "<a:blip rel:embed=\"" + blipIdPng + "\">" +
                        "<a:extLst>" +
                        "<a:ext uri=\"{28A0092B-C50C-407E-A947-70E740481C1C}\">" +
                        "<a14:useLocalDpi\n" +
                        "xmlns:a14=\"http://schemas.microsoft.com/office/drawing/2010/main\" val=\"0\"/>" +
                        "</a:ext>" +
                        "<a:ext uri=\"{96DAC541-7B7A-43D3-8B79-37D633B846F1}\">" +
                        "<asvg:svgBlip\n" +
                        "xmlns:asvg=\"http://schemas.microsoft.com/office/drawing/2016/SVG/main\" rel:embed=\"" + blipId + "\"/>" +
                        "</a:ext>" +
                        "</a:extLst>" +
                        "</a:blip>" +
                        "<a:stretch>" +
                        "<a:fillRect/>" +
                        "</a:stretch>" +
                        "</pic:blipFill>" +
                        "<pic:spPr>" +
                        "<a:xfrm>" +
                        "<a:off x=\"0\" y=\"0\"/>" +
                        "<a:ext cx=\"" + widthString + "\" cy=\"" + heightString + "\"/>" +
                        "</a:xfrm>" +
                        "<a:prstGeom prst=\"rect\">" +
                        "<a:avLst/>" +
                        "</a:prstGeom>" +
                        "</pic:spPr>" +
                        "</pic:pic>" +
                        "</a:graphicData>" +
                        "</a:graphic>" +
                        "</wp:inline>" +
                        "</w:drawing>"

                XmlToken xmlToken = XmlToken.Factory.parse(xmlSvgStructure)
                ctr.set(xmlToken)
            }
        }
    }

    enum SubprocessOwnerType {
        ORGANIZATIONAL_UNIT,
        GROUP,
    }

    private static final Map<String, SubprocessOwnerType> subprocessOwnerTypeMap
    static {
        subprocessOwnerTypeMap = new HashMap<>()
        subprocessOwnerTypeMap.put(ORGANIZATIONAL_UNIT_OBJECT_TYPE_ID, SubprocessOwnerType.ORGANIZATIONAL_UNIT)
        subprocessOwnerTypeMap.put(GROUP_OBJECT_TYPE_ID, SubprocessOwnerType.GROUP)
    }

    private class CommonObjectInfo {
        ObjectElement object
        String name

        CommonObjectInfo(ObjectElement object, boolean onlyShortName = false) {
            this.object = object
            this.name = getName(object.getObjectDefinition(), onlyShortName)
        }

        CommonObjectInfo(Model model, boolean onlyShortName = false) {
            this.object = null
            this.name = getName(model, onlyShortName)
        }
    }

    private class CommonFunctionInfo {
        CommonObjectInfo function
        String code
        String requirements

        CommonFunctionInfo(ObjectElement function) {
            this.function = new CommonObjectInfo(function)
            ObjectDefinition objectDefinition = function.getObjectDefinition()
            this.code = getAttributeValue(objectDefinition, DATA_ELEMENT_CODE_ATTR_ID)
            this.requirements = getAttributeValue(objectDefinition, DESCRIPTION_DEFINITION_ATTR_ID)
        }

        CommonFunctionInfo(Model model) {
            this.function = new CommonObjectInfo(model)
            this.code = getAttributeValue(model, DATA_ELEMENT_CODE_ATTR_ID)
            this.requirements = getAttributeValue(model, DESCRIPTION_DEFINITION_ATTR_ID)
        }
    }

    private class PositionInfo {
        CommonObjectInfo position
        CommonObjectInfo organizationalUnit

        PositionInfo(ObjectElement position) {
            this.position = new CommonObjectInfo(position)
            defineOrganizationalUnit()
        }

        private void defineOrganizationalUnit() {
            List<ObjectElement> positionInstances = position.object.getObjectDefinition().getInstances()
            for (instance in positionInstances) {
                ObjectElement organizationalUnitObject = instance.getEnterEdges()
                        .findAll { Edge e -> e.getEdgeTypeId() == ORGANIZATIONAL_UNIT_W_POSITION_EDGE_TYPE_ID }
                        .collect { Edge e -> e.getSource() as ObjectElement }
                        .unique(Comparator.comparing { ObjectElement oE -> oE.getId() })
                        .stream()
                        .findFirst()
                        .orElse(null)

                if (organizationalUnitObject) {
                    organizationalUnit = new CommonObjectInfo(organizationalUnitObject)
                    break
                }
            }
        }
    }

    private class BusinessRoleInfo {
        CommonObjectInfo businessRole
        List<PositionInfo> positions = []

        BusinessRoleInfo(ObjectElement businessRole) {
            this.businessRole = new CommonObjectInfo(businessRole)
            definePositions()
        }

        private void definePositions() {
            List<ObjectElement> businessRoleInstances = businessRole.object.getObjectDefinition().getInstances()
            for (instance in businessRoleInstances) {
                List<ObjectElement> positionObjects = instance.getEnterEdges()
                        .findAll { Edge e -> e.getEdgeTypeId() == POSITION_W_BUSINESS_ROLE_EDGE_TYPE_ID }
                        .collect { Edge e -> e.getSource() as ObjectElement }
                        .unique(Comparator.comparing { ObjectElement oE -> oE.getId() })
                positions.addAll(positionObjects.collect { ObjectElement positionObject -> new PositionInfo(positionObject) })
            }
        }
    }

    private class SubprocessOwnerInfo {
        CommonObjectInfo owner
        SubprocessOwnerType type
        String leadershipPosition = null

        SubprocessOwnerInfo(ObjectElement owner, SubprocessOwnerType ownerType) {
            this.owner = new CommonObjectInfo(owner)
            this.type = ownerType

            if (type == SubprocessOwnerType.ORGANIZATIONAL_UNIT) {
                defineLeadershipPosition()
            }
        }

        private void defineLeadershipPosition() {
            ObjectDefinition ownerObjectDefinition = owner.object.getObjectDefinition()
            Model ownerModel = ownerObjectDefinition
                    .getDecompositions(ORGANIZATION_STRUCTURE_MODEL_TYPE_ID)
                    .stream()
                    .findFirst()
                    .orElse(null)

            if (ownerModel == null) {
                return
            }

            ObjectElement ownerModelObject = ownerModel.findObjectInstances(ownerObjectDefinition)
                    .stream()
                    .findFirst()
                    .orElse(null)

            if (ownerModelObject == null) {
                return
            }

            ObjectElement leadershipPositionObject = ownerModelObject.getEnterEdges()
                    .find { Edge e -> e.getEdgeTypeId() == LEADERSHIP_POSITION_W_OWNER_EDGE_TYPE_ID }
                    .getSource() as ObjectElement
            this.leadershipPosition = getName(leadershipPositionObject.getObjectDefinition())
        }
    }

    private class PerformerInfo {
        CommonObjectInfo performer
        String action

        PerformerInfo(ObjectElement performer, Edge edge) {
            this.performer = new CommonObjectInfo(performer)
            this.action = edge.getEdgeType().name
        }
    }

    private class DocumentInfo {
        CommonObjectInfo document
        String type
        List<CommonObjectInfo> statuses = []

        DocumentInfo(ObjectElement document) {
            this.document = new CommonObjectInfo(document)
            this.type = document.getSymbol().name
            findStatuses()
        }

        private void findStatuses() {
            List<ObjectElement> statusObjects = document.object.getExitEdges()
                    .findAll { Edge e -> e.getEdgeTypeId() == DOCUMENT_W_STATUS_EDGE_TYPE_ID }
                    .collect { Edge e -> e.getTarget() as ObjectElement }
                    .findAll { ObjectElement oE -> oE.getSymbolId() == STATUS_SYMBOL_ID }
                    .unique(Comparator.comparing { ObjectElement oE -> oE.getObjectDefinitionId() })
                    .sort { ObjectElement oE1, ObjectElement oE2 -> ModelUtils.getElementsCoordinatesComparator().compare(oE1, oE2) }
            statuses = statusObjects.collect { ObjectElement statusObject -> new CommonObjectInfo(statusObject) }
        }
    }

    private class NormativeDocumentInfo {
        DocumentInfo document
        String requisites

        NormativeDocumentInfo(ObjectElement document) {
            this.document = new DocumentInfo(document)
            this.requisites = getAttributeValue(document.getObjectDefinition(), DESCRIPTION_DEFINITION_ATTR_ID)
        }
    }

    private class DocumentCollectionInfo {
        DocumentInfo collection
        Model model
        List<DocumentInfo> containedDocuments = []

        DocumentCollectionInfo(ObjectElement collection, Model model) {
            this.collection = new DocumentInfo(collection)
            this.model = model
        }

        DocumentCollectionInfo(DocumentInfo collection, Model model) {
            this.collection = collection
            this.model = model
        }

        void findContainedDocuments() {
            ObjectElement collectionObjectOnModel = model.findObjectInstances(collection.document.object.getObjectDefinition())
                    .stream()
                    .findFirst()
                    .orElse(null)

            if (collectionObjectOnModel == null) {
                return
            }

            String modelTypeId = model.getModelTypeId()
            List<ObjectElement> containedDocumentObjects = []

            if (modelTypeId == IEF_DATA_MODEL_TYPE_ID) {
                containedDocumentObjects.addAll(collectionObjectOnModel.getExitEdges()
                        .findAll { Edge e -> e.getEdgeTypeId() in CLUSTER_GROUP_W_CLUSTER_DATA_MODEL_EDGE_TYPE_IDS }
                        .collect { Edge e -> e.getTarget() as ObjectElement }
                        .findAll { ObjectElement oE -> oE.getObjectDefinition().getObjectTypeId() == CLUSTER_DATA_MODEL_OBJECT_TYPE_ID }
                        .unique(Comparator.comparing { ObjectElement oE -> oE.getObjectDefinitionId() })
                        .sort { ObjectElement oE1, ObjectElement oE2 -> ModelUtils.getElementsCoordinatesComparator().compare(oE1, oE2) }
                )
            }

            if (modelTypeId == INFORMATION_CARRIER_MODEL_TYPE_ID) {
                containedDocumentObjects.addAll(collectionObjectOnModel.getExitEdges()
                        .findAll { Edge e -> e.getEdgeTypeId() in DOCUMENT_COLLECTION_W_DOCUMENT_EDGE_TYPE_IDS }
                        .collect { Edge e -> e.getTarget() as ObjectElement }
                        .findAll { ObjectElement oE -> oE.getObjectDefinition().getObjectTypeId() == INFORMATION_CARRIER_OBJECT_TYPE_ID }
                        .unique(Comparator.comparing { ObjectElement oE -> oE.getObjectDefinitionId() })
                        .sort { ObjectElement oE1, ObjectElement oE2 -> ModelUtils.getElementsCoordinatesComparator().compare(oE1, oE2) }
                )
            }

            containedDocuments = containedDocumentObjects.collect { ObjectElement containedDocumentObject -> new DocumentInfo(containedDocumentObject) }
        }
    }

    private class FileInfo {
        String name
        FileNodeDTO content = null

        FileInfo(String name, FileNodeDTO content) {
            this.name = name
            this.content = content
        }
    }

    private class SubprocessDescription {
        private class ExternalProcessDescription {
            CommonFunctionInfo externalProcess
            List<CommonObjectInfo> flows

            ExternalProcessDescription(CommonFunctionInfo externalProcess, List<CommonObjectInfo> flows) {
                this.externalProcess = externalProcess
                this.flows = flows
            }
        }

        CommonFunctionInfo subprocess
        int detailLevel

        CommonFunctionInfo parentProcess = null
        List<SubprocessOwnerInfo> owners = []
        List<CommonObjectInfo> goals = []
        List<InputFlowDescription> externalProcessInputFlowDescriptions = []
        List<OutputFlowDescription> externalProcessOutputFlowDescriptions = []
        Model processSelectionModel = null
        List<ScenarioDescription> scenarios = []

        List<ExternalProcessDescription> completedExternalProcessesWithInputFlows = []
        List<ExternalProcessDescription> completedExternalProcessesWithOutputFlows = []
        List<BusinessRoleInfo> completedBusinessRoles = []
        List<EPCDescription> analyzedEPC = []
        List<DocumentCollectionInfo> completedDocumentCollections = []
        List<NormativeDocumentInfo> completedNormativeDocuments = []

        SubprocessDescription(ObjectElement subprocess, int detailLevel) {
            this.subprocess = new CommonFunctionInfo(subprocess)
            this.detailLevel = detailLevel
        }

        void defineParentProcess() {
            List<ObjectDefinition> parentObjects = subprocess.function.object.model.parentObjects

            ObjectDefinition parentObject = null
            Model parentModel = null
            for (object in parentObjects) {
                if (parentObject) {
                    break
                }

                List<Model> parentModels = object.getParentModels()
                for (model in parentModels) {
                    if (model.getId() == FIRST_LEVEL_MODEL_ID) {
                        parentObject = object
                        parentModel = model
                        break
                    }
                }
            }

            if (parentObject == null) {
                return
            }

            ObjectElement parentElement = parentModel.findObjectInstances(parentObject)
                    .stream()
                    .findFirst()
                    .orElse(null)
            this.parentProcess = new CommonFunctionInfo(parentElement)
        }

        void findOwners() {
            List<ObjectElement> ownerObjects = subprocess.function.object.getEnterEdges()
                    .findAll { Edge e -> e.getEdgeTypeId() in OWNER_W_SUBPROCESS_EDGE_TYPE_IDS }
                    .collect { Edge e -> e.getSource() as ObjectElement }
                    .unique(Comparator.comparing { ObjectElement oE -> oE.getId() })
                    .sort { ObjectElement oE1, ObjectElement oE2 -> ModelUtils.getElementsCoordinatesComparator().compare(oE1, oE2) }
            owners = ownerObjects.collect { ObjectElement ownerObject ->
                new SubprocessOwnerInfo(ownerObject, subprocessOwnerTypeMap.get(ownerObject.getObjectDefinition().getObjectTypeId()))
            }
        }

        void defineGoals() {
            Model functionAllocationModel = subprocess.function.object.getObjectDefinition()
                    .getDecompositions(FUNCTION_ALLOCATION_MODEL_TYPE_ID)
                    .stream()
                    .findFirst()
                    .orElse(null)

            if (functionAllocationModel == null) {
                return
            }

            List<ObjectElement> goalObjects = functionAllocationModel.findObjectsByType(GOAL_OBJECT_TYPE_ID)
            goals = goalObjects.collect { ObjectElement goalObject -> new CommonObjectInfo(goalObject) }
        }

        void findExternalProcessInputFlows() {
            List<ObjectElement> allFlowObjects = subprocess.function.object.model.findObjectsByType(FLOW_OBJECT_TYPE_ID)

            List<ObjectElement> inputFlowObjects = subprocess.function.object.getEnterEdges()
                    .findAll { Edge e -> e.getEdgeTypeId() == INPUT_FLOW_W_SUBPROCESS_EDGE_TYPE_ID }
                    .collect { Edge e -> e.getSource() as ObjectElement }
                    .findAll { ObjectElement oE -> oE.getObjectDefinition().getObjectTypeId() == FLOW_OBJECT_TYPE_ID }
                    .unique(Comparator.comparing { ObjectElement oE -> oE.getId() })
                    .sort { ObjectElement oE1, ObjectElement oE2 -> ModelUtils.getElementsCoordinatesComparator().compare(oE1, oE2) }

            inputFlowObjects.each { ObjectElement currentFlowObject ->
                List<ObjectElement> externalSupplierObjects = currentFlowObject.getEnterEdges()
                        .findAll { Edge e -> e.getEdgeTypeId() == SUPPLIER_W_INPUT_FLOW_EDGE_TYPE_ID }
                        .collect { Edge e -> e.getSource() as ObjectElement }
                        .findAll { ObjectElement oE -> oE.getSymbolId() == EXTERNAL_PROCESS_SYMBOL_ID }
                        .unique(Comparator.comparing { ObjectElement oE -> oE.getId() })
                        .sort { ObjectElement oE1, ObjectElement oE2 -> ModelUtils.getElementsCoordinatesComparator().compare(oE1, oE2) }

                List<ObjectElement> additionalExternalSupplierObjects = findAdditionalExternalSupplierObjects(currentFlowObject, allFlowObjects)
                externalSupplierObjects.addAll(additionalExternalSupplierObjects)
                externalSupplierObjects = externalSupplierObjects
                        .unique(Comparator.comparing { ObjectElement oE -> oE.getObjectDefinitionId() })
                        .sort { ObjectElement oE1, ObjectElement oE2 -> ModelUtils.getElementsCoordinatesComparator().compare(oE1, oE2) }

                if (externalSupplierObjects) {
                    externalProcessInputFlowDescriptions.add(new InputFlowDescription(currentFlowObject, externalSupplierObjects))
                }
            }
        }

        private List<ObjectElement> findAdditionalExternalSupplierObjects(ObjectElement currentFlowObject, List<ObjectElement> allFlowObjects) {
            String currentFlowObjectDefinitionId = currentFlowObject.getObjectDefinitionId()
            List<ObjectElement> currentFlowObjects = allFlowObjects
                    .findAll { ObjectElement flowObject -> flowObject.getObjectDefinitionId() == currentFlowObjectDefinitionId }

            List<ObjectElement> additionalExternalSupplierObjects = []
            for (flowObject in currentFlowObjects) {
                if (flowObject.getId() == currentFlowObject.getId()) {
                    continue
                }

                List<ObjectElement> foundedObjects = flowObject.getEnterEdges()
                        .findAll { Edge e -> e.getEdgeTypeId() == SUBPROCESS_W_OUTPUT_FLOW_EDGE_TYPE_ID }
                        .collect { Edge e -> e.getSource() as ObjectElement }
                        .findAll { ObjectElement oE -> oE.getSymbolId() == EXTERNAL_PROCESS_SYMBOL_ID }
                        .unique(Comparator.comparing { ObjectElement oE -> oE.getId() })
                        .sort { ObjectElement oE1, ObjectElement oE2 -> ModelUtils.getElementsCoordinatesComparator().compare(oE1, oE2) }
                additionalExternalSupplierObjects.addAll(foundedObjects)
            }
            return additionalExternalSupplierObjects
        }

        void findExternalProcessOutputFlows() {
            List<ObjectElement> outputFlowObjects = subprocess.function.object.getExitEdges()
                    .findAll { Edge e -> e.getEdgeTypeId() == SUBPROCESS_W_OUTPUT_FLOW_EDGE_TYPE_ID }
                    .collect { Edge e -> e.getTarget() as ObjectElement }
                    .findAll { ObjectElement oE -> oE.getObjectDefinition().getObjectTypeId() == FLOW_OBJECT_TYPE_ID }
                    .unique(Comparator.comparing { ObjectElement oE -> oE.getId() })
                    .sort { ObjectElement oE1, ObjectElement oE2 -> ModelUtils.getElementsCoordinatesComparator().compare(oE1, oE2) }

            outputFlowObjects.each { ObjectElement currentFlowObject ->
                List<ObjectElement> externalCustomerObjects = currentFlowObject.getExitEdges()
                        .findAll { Edge e -> e.getEdgeTypeId() == OUTPUT_FLOW_W_CUSTOMER_EDGE_TYPE_ID }
                        .collect { Edge e -> e.getTarget() as ObjectElement }
                        .findAll { ObjectElement oE -> oE.getSymbolId() == EXTERNAL_PROCESS_SYMBOL_ID }
                        .unique(Comparator.comparing { ObjectElement oE -> oE.getId() })
                        .sort { ObjectElement oE1, ObjectElement oE2 -> ModelUtils.getElementsCoordinatesComparator().compare(oE1, oE2) }

                if (externalCustomerObjects) {
                    externalProcessOutputFlowDescriptions.add(new OutputFlowDescription(currentFlowObject, externalCustomerObjects))
                }
            }
        }

        void completeExternalProcessesWithInputFlows() {
            externalProcessInputFlowDescriptions.each { InputFlowDescription inputFlowDescription ->
                addExternalProcessesWithFlow(inputFlowDescription.inputFlow, inputFlowDescription.suppliers, completedExternalProcessesWithInputFlows)
            }
        }

        void completeExternalProcessesWithOutputFlows() {
            externalProcessOutputFlowDescriptions.each { OutputFlowDescription outputFlowDescription ->
                addExternalProcessesWithFlow(outputFlowDescription.outputFlow, outputFlowDescription.customers, completedExternalProcessesWithOutputFlows)
            }
        }

        private void addExternalProcessesWithFlow(CommonObjectInfo flow, List<CommonFunctionInfo> externalProcesses, List<ExternalProcessDescription> completedExternalProcessesWithFlows) {
            for (process in externalProcesses) {
                List<String> completedProcessObjectDefinitionIds = completedExternalProcessesWithFlows.collect { ExternalProcessDescription ePWF -> ePWF.externalProcess.function.object.getObjectDefinitionId() }
                String processObjectDefinitionId = process.function.object.getObjectDefinitionId()

                if (processObjectDefinitionId in completedProcessObjectDefinitionIds) {
                    ExternalProcessDescription processDescription = completedExternalProcessesWithFlows
                            .find { ExternalProcessDescription ePWF -> ePWF.externalProcess.function.object.getObjectDefinitionId() == processObjectDefinitionId }

                    List<String> completedFlowObjectDefinitionIds = processDescription.flows.collect { CommonObjectInfo f -> f.object.getObjectDefinitionId() }
                    if (flow.object.getObjectDefinitionId() in completedFlowObjectDefinitionIds) {
                        continue
                    }

                    processDescription.flows.add(flow)
                } else {
                    completedExternalProcessesWithFlows.add(new ExternalProcessDescription(process, [flow]))
                }
            }
        }

        void defineProcessSelectionModel() {
            processSelectionModel = subprocess.function.object.getObjectDefinition()
                    .getDecompositions(PROCESS_SELECTION_MODEL_TYPE_ID)
                    .stream()
                    .findFirst()
                    .orElse(null)
        }

        void defineScenarios() {
            if (processSelectionModel) {
                defineScenariosViaProcessSelectionModel()
                return
            }

            Model scenarioModel = getEPCModel(subprocess.function.object)
            if (scenarioModel) {
                scenarios.add(new ScenarioDescription(scenarioModel))
            }
        }

        private void defineScenariosViaProcessSelectionModel() {
            List<ObjectElement> scenarioObjects = processSelectionModel.getObjects()
                    .findAll { ObjectElement oE -> oE.getSymbolId() in SCENARIO_SYMBOL_IDS }
                    .unique(Comparator.comparing { ObjectElement oE -> oE.getId() })

            Map<String, ObjectElement> parentScenarioObjectMap = new HashMap<>()
            scenarioObjects.each { ObjectElement scenarioObject ->
                String parent = scenarioObject._getDiagramElement().parent
                parentScenarioObjectMap.put(parent, scenarioObject)
            }

            ModelNode modelNode = (ModelNode) processSelectionModel._getNode().asNodeSubtype()
            List<GridCell> modelCells = modelNode.layout.cells
            List<GridHeader> columns = modelNode.layout.columns

            Map<Integer, String> columnParentMap = new TreeMap<>()
            for (parent in parentScenarioObjectMap.keySet()) {
                GridCell cell = modelCells.find { GridCell c -> c.id == parent }
                String columnId = cell.columnId

                columns.eachWithIndex { GridHeader column, int number ->
                    if (column.id == columnId) {
                        columnParentMap.put(number, parent)
                        return
                    }
                }
            }

            for (parent in columnParentMap.values()) {
                ObjectElement scenarioObject = parentScenarioObjectMap.get(parent)
                Model scenarioModel = getEPCModel(scenarioObject)

                if (scenarioModel == null) {
                    return
                }

                scenarios.add(new ScenarioDescription(scenarioModel, scenarioObject))
            }
        }

        void defineProcedures() {
            scenarios.each { ScenarioDescription scenarioDescription -> scenarioDescription.defineProcedures() }
        }

        void defineBusinessRoles() {
            scenarios.each { ScenarioDescription scenarioDescription -> scenarioDescription.defineBusinessRoles() }
        }

        void completeBusinessRoles() {
            scenarios.each { ScenarioDescription scenarioDescription ->
                completedBusinessRoles.addAll(scenarioDescription.getAllBusinessRoles())
            }
            completedBusinessRoles = completedBusinessRoles
                    .unique(Comparator.comparing { BusinessRoleInfo bRI -> bRI.businessRole.object.getObjectDefinitionId() })
        }

        void buildResponsibilityScenariosMatrix() {
            scenarios.each { ScenarioDescription scenarioDescription -> scenarioDescription.buildResponsibilityMatrix() }
        }

        void identifyAnalyzedEPC() {
            if (detailLevel == 3) {
                for (scenario in scenarios) {
                    analyzedEPC.add(scenario.scenario)
                }
            }

            if (detailLevel == 4) {
                for (scenario in scenarios) {
                    for (procedure in scenario.procedures) {
                        analyzedEPC.add(procedure.procedure)
                    }
                }
            }
        }

        void defineNormativeDocuments() {
            analyzedEPC.each { EPCDescription epcDescription -> epcDescription.findNormativeDocuments() }
        }

        void completeNormativeDocuments() {
            analyzedEPC.each { EPCDescription epcDescription ->
                completedNormativeDocuments.addAll(epcDescription.normativeDocuments)
            }

            for (documentCollection in completedDocumentCollections) {
                for (containedDocument in documentCollection.containedDocuments) {
                    if (containedDocument.document.object.getSymbolId() == NORMATIVE_DOCUMENT_SYMBOL_ID) {
                        completedNormativeDocuments.add(new NormativeDocumentInfo(containedDocument.document.object))
                    }
                }
            }

            completedNormativeDocuments = completedNormativeDocuments
                    .unique(Comparator.comparing { NormativeDocumentInfo nDI -> nDI.document.document.object.getObjectDefinitionId() })
        }

        void defineDocumentCollections() {
            analyzedEPC.each { EPCDescription epcDescription -> epcDescription.findDocumentCollections() }
        }

        void completeDocumentCollections() {
            for (epcDescription in analyzedEPC) {
                for (documentCollection in epcDescription.documentCollections) {
                    completedDocumentCollections.add(documentCollection)
                }
            }

            completedDocumentCollections = completedDocumentCollections
                    .unique(Comparator.comparing { DocumentCollectionInfo dCO -> dCO.collection.document.object.getObjectDefinitionId() })

            List<DocumentCollectionInfo> foundedDocumentCollections = completedDocumentCollections
            while (foundedDocumentCollections) {
                List<DocumentCollectionInfo> unparsedDocumentCollections = foundedDocumentCollections
                unparsedDocumentCollections.each { DocumentCollectionInfo dCO -> dCO.findContainedDocuments() }
                foundedDocumentCollections = parseDocumentCollections(unparsedDocumentCollections)
                completedDocumentCollections.addAll(foundedDocumentCollections)
            }
        }

        private List<DocumentCollectionInfo> parseDocumentCollections(List<DocumentCollectionInfo> unparsedDocumentCollections) {
            List<String> completedDocumentCollectionObjectDefinitionIds = completedDocumentCollections.collect { DocumentCollectionInfo dCO -> dCO.collection.document.object.getObjectDefinitionId() }
            List<DocumentCollectionInfo> foundedDocumentCollections = []
            for (unparsedDocumentCollection in unparsedDocumentCollections) {
                for (containedDocument in unparsedDocumentCollection.containedDocuments) {
                    Model containedDocumentModel = EPCDescription.findDocumentCollectionModel(containedDocument.document.object)
                    boolean containedDocumentAlreadyInCompletedCollections = containedDocument.document.object.getObjectDefinitionId() in completedDocumentCollectionObjectDefinitionIds

                    if (containedDocumentModel && !containedDocumentAlreadyInCompletedCollections) {
                        foundedDocumentCollections.add(new DocumentCollectionInfo(containedDocument, containedDocumentModel))
                    }
                }
            }
            return foundedDocumentCollections
        }

        void analyzeEPCModels() {
            analyzedEPC.each { EPCDescription epc -> analyzeEPCModel(epc) }
        }

        private analyzeEPCModel(EPCDescription epc) {
            epc.findFunctions()
            epc.analyzeFunctions()
        }
    }

    private class ScenarioDescription {
        EPCDescription scenario
        List<ProcedureDescription> procedures = []
        Map<String, List<String>> responsibilityMatrix = new TreeMap<>()

        ScenarioDescription(Model model, ObjectElement functionObject) {
            this.scenario = new EPCDescription(model, functionObject)
        }

        ScenarioDescription(Model model) {
            this.scenario = new EPCDescription(model)
        }

        void defineProcedures() {
            List<ObjectElement> procedureObjects = scenario.model.findObjectsByType(FUNCTION_OBJECT_TYPE_ID)
                    .findAll { ObjectElement functionObject -> !(functionObject.getSymbolId() in EXCLUDED_FUNCTION_SYMBOL_IDS) }
                    .unique(Comparator.comparing { ObjectElement oE -> oE.getId() })
                    .sort { ObjectElement oE1, ObjectElement oE2 -> ModelUtils.getElementsCoordinatesComparator().compare(oE1, oE2) }

            procedureObjects.each { ObjectElement procedureObject ->
                Model procedureModel = getEPCModel(procedureObject)

                if (procedureModel == null) {
                    return
                }

                procedures.add(new ProcedureDescription(procedureModel, procedureObject))
            }
        }

        void defineBusinessRoles() {
            procedures.each { ProcedureDescription procedureDescription -> procedureDescription.findBusinessRoles() }
        }

        List<BusinessRoleInfo> getAllBusinessRoles() {
            List<BusinessRoleInfo> allBusinessRoles = []
            procedures.each { ProcedureDescription procedureDescription ->
                allBusinessRoles.addAll(procedureDescription.businessRoles)
            }
            return allBusinessRoles.unique(Comparator.comparing { BusinessRoleInfo bRI -> bRI.businessRole.object.getObjectDefinitionId() })
        }

        void buildResponsibilityMatrix() {
            for (procedure in procedures) {
                String procedureName = procedure.procedure.functionInfo.function.name
                List<String> procedureBusinessRoleNames = procedure.businessRoles.collect { BusinessRoleInfo businessRole -> businessRole.businessRole.name }
                responsibilityMatrix.put(procedureName, procedureBusinessRoleNames)
            }
        }
    }

    private class ProcedureDescription {
        EPCDescription procedure
        List<BusinessRoleInfo> businessRoles = []

        ProcedureDescription(Model model, ObjectElement functionObject) {
            this.procedure = new EPCDescription(model, functionObject)
        }

        void findBusinessRoles() {
            List<ObjectElement> businessRoleObjects = procedure.model.findObjectsByType(BUSINESS_ROLE_OBJECT_TYPE_ID)
                    .unique(Comparator.comparing { ObjectElement oE -> oE.getObjectDefinitionId() })
                    .sort { ObjectElement oE1, ObjectElement oE2 -> ModelUtils.getElementsCoordinatesComparator().compare(oE1, oE2) }
            businessRoles = businessRoleObjects.collect { ObjectElement businessRoleObject -> new BusinessRoleInfo(businessRoleObject) }
        }
    }

    private class EPCDescription {
        CommonFunctionInfo functionInfo
        Model model
        List<NormativeDocumentInfo> normativeDocuments = []
        List<DocumentCollectionInfo> documentCollections = []
        List<EPCFunctionDescription> epcFunctions = []

        EPCDescription(Model model, ObjectElement functionObject) {
            this.functionInfo = new CommonFunctionInfo(functionObject)
            this.model = model
        }

        EPCDescription(Model model) {
            this.functionInfo = new CommonFunctionInfo(model)
            this.model = model
        }

        void findNormativeDocuments() {
            List<ObjectElement> normativeDocumentObjects = model.getObjects()
                    .findAll { ObjectElement oE -> oE.getSymbolId() == NORMATIVE_DOCUMENT_SYMBOL_ID }
                    .unique(Comparator.comparing { ObjectElement oE -> oE.getObjectDefinitionId() })
            normativeDocuments = normativeDocumentObjects.collect { ObjectElement normativeDocumentObject -> new NormativeDocumentInfo(normativeDocumentObject) }
        }

        void findDocumentCollections() {
            List<ObjectElement> documentObjects = model.getObjects()
                    .findAll { ObjectElement oE -> oE.getObjectDefinition().getObjectTypeId() in DOCUMENT_OBJECT_TYPE_IDS }
                    .unique(Comparator.comparing { ObjectElement oE -> oE.getObjectDefinitionId() })
                    .sort { ObjectElement oE1, ObjectElement oE2 -> ModelUtils.getElementsCoordinatesComparator().compare(oE1, oE2) }

            documentObjects.each { ObjectElement documentObject ->
                Model documentCollectionModel = findDocumentCollectionModel(documentObject)

                if (documentCollectionModel) {
                    documentCollections.add(new DocumentCollectionInfo(documentObject, documentCollectionModel))
                }
            }
        }

        static Model findDocumentCollectionModel(ObjectElement documentCollectionObject) {
            List<Model> documentCollectionObjectModels = documentCollectionObject.getDecompositions()
                    .findAll { TreeNode tN -> tN.isModel() } as List<Model>
            return documentCollectionObjectModels
                    .findAll { Model m -> m.getModelTypeId() in DOCUMENT_COLLECTION_MODEL_TYPE_IDS }
                    .stream()
                    .findFirst()
                    .orElse(null)
        }

        void findFunctions() {
            List<ObjectElement> epcFunctionObjects = model.findObjectsByType(FUNCTION_OBJECT_TYPE_ID)
                    .findAll { ObjectElement epcFunctionObject -> !(epcFunctionObject.getSymbolId() in EXCLUDED_FUNCTION_SYMBOL_IDS) }
                    .unique(Comparator.comparing { ObjectElement oE -> oE.getId() })
                    .sort { ObjectElement oE1, ObjectElement oE2 -> ModelUtils.getElementsCoordinatesComparator().compare(oE1, oE2) }

            int number = 1
            for (epcFunctionObject in epcFunctionObjects) {
                epcFunctions.add(new EPCFunctionDescription(epcFunctionObject, number))
                number += 1
            }
        }

        void analyzeFunctions() {
            epcFunctions.each { EPCFunctionDescription epcFunction -> analyzeFunction(epcFunction) }
        }

        private void analyzeFunction(EPCFunctionDescription epcFunction) {
            epcFunction.findInputDocuments()
            epcFunction.findInputEvents()
            epcFunction.findOutputDocuments()
            epcFunction.findOutputEvents()
            epcFunction.findPerformers()
            epcFunction.findInformationSystems()
            epcFunction.findChildFunctions(epcFunctions)
        }
    }

    private class EPCFunctionDescription {
        CommonFunctionInfo function
        int number
        String duration

        List<DocumentInfo> inputDocuments = []
        List<CommonObjectInfo> inputEvents = []
        List<DocumentInfo> outputDocuments = []
        List<CommonObjectInfo> outputEvents = []
        List<PerformerInfo> performers = []
        List<CommonObjectInfo> informationSystems = []

        List<EPCFunctionDescription> childEPCFunctions = []
        List<CommonFunctionInfo> childExternalFunctions = []

        EPCFunctionDescription(ObjectElement function, int number) {
            this.function = new CommonFunctionInfo(function)
            this.number = number
            this.duration = getAttributeValue(function.getObjectDefinition(), AVERAGE_EXECUTION_TIME_ATTR_ID)
        }

        void findInputDocuments() {
            List<ObjectElement> documentObjects = function.function.object.getEnterEdges()
                    .findAll { Edge e -> e.getEdgeTypeId() in DOCUMENT_W_EPC_FUNCTION_EDGE_TYPE_IDS }
                    .collect { Edge e -> e.getSource() as ObjectElement }
                    .findAll { ObjectElement oE -> oE.getObjectDefinition().getObjectTypeId() in DOCUMENT_OBJECT_TYPE_IDS }
                    .sort { ObjectElement oE1, ObjectElement oE2 -> ModelUtils.getElementsCoordinatesComparator().compare(oE1, oE2) }
            inputDocuments = documentObjects.collect { ObjectElement documentObject -> new DocumentInfo(documentObject) }
        }

        void findInputEvents() {
            List<ObjectElement> eventObjects = function.function.object.getEnterEdges()
                    .findAll { Edge e -> e.getEdgeTypeId() == EVENT_W_EPC_FUNCTION_EDGE_TYPE_ID }
                    .collect { Edge e -> e.getSource() as ObjectElement }
                    .findAll { ObjectElement oE -> oE.getObjectDefinition().getObjectTypeId() == EVENT_OBJECT_TYPE_ID }
                    .unique(Comparator.comparing { ObjectElement oE -> oE.getId() })

            List<ObjectElement> operators = function.function.object.getEnterEdges()
                    .findAll { Edge e -> e.getEdgeTypeId() == OPERATOR_W_EPC_FUNCTION_EDGE_TYPE_ID }
                    .collect { Edge e -> e.getSource() as ObjectElement }
                    .findAll { ObjectElement oE -> oE.getObjectDefinition().getObjectTypeId() == RULE_OBJECT_TYPE_ID }
                    .unique(Comparator.comparing { ObjectElement oE -> oE.getId() })

            while (operators) {
                List<ObjectElement> unparsedOperators = operators
                operators = []

                unparsedOperators.each { ObjectElement unparsedOperator ->
                    eventObjects.addAll(findInputEventsViaOperator(unparsedOperator))
                    operators.addAll(findInputOperatorsViaOperator(unparsedOperator))
                }
            }

            eventObjects = eventObjects
                    .unique(Comparator.comparing { ObjectElement oE -> oE.getObjectDefinitionId() })
                    .sort { ObjectElement oE1, ObjectElement oE2 -> ModelUtils.getElementsCoordinatesComparator().compare(oE1, oE2) }
            inputEvents = eventObjects.collect { ObjectElement eventObject -> new CommonObjectInfo(eventObject) }
        }

        private List<ObjectElement> findInputEventsViaOperator(ObjectElement operator) {
            List<ObjectElement> eventObjects = operator.getEnterEdges()
                    .findAll { Edge e -> e.getEdgeTypeId() == EVENT_W_OPERATOR_EDGE_TYPE_ID }
                    .collect { Edge e -> e.getSource() as ObjectElement }
                    .findAll { ObjectElement oE -> oE.getObjectDefinition().getObjectTypeId() == EVENT_OBJECT_TYPE_ID }
                    .unique(Comparator.comparing { ObjectElement oE -> oE.getId() })
            return eventObjects
        }

        private List<ObjectElement> findInputOperatorsViaOperator(ObjectElement operator) {
            List<ObjectElement> operators = operator.getEnterEdges()
                    .findAll { Edge e -> e.getEdgeTypeId() in OPERATOR_W_OPERATOR_EDGE_TYPE_IDS }
                    .collect { Edge e -> e.getSource() as ObjectElement }
                    .findAll { ObjectElement oE -> oE.getObjectDefinition().getObjectTypeId() == RULE_OBJECT_TYPE_ID }
                    .unique(Comparator.comparing { ObjectElement oE -> oE.getId() })
            return operators
        }

        void findOutputDocuments() {
            List<ObjectElement> documentObjects = function.function.object.getExitEdges()
                    .findAll { Edge e -> e.getEdgeTypeId() in EPC_FUNCTION_W_DOCUMENT_EDGE_TYPE_IDS }
                    .collect { Edge e -> e.getTarget() as ObjectElement }
                    .findAll { ObjectElement oE -> oE.getObjectDefinition().getObjectTypeId() in DOCUMENT_OBJECT_TYPE_IDS }
                    .sort { ObjectElement oE1, ObjectElement oE2 -> ModelUtils.getElementsCoordinatesComparator().compare(oE1, oE2) }
            outputDocuments = documentObjects.collect { ObjectElement documentObject -> new DocumentInfo(documentObject) }
        }

        void findOutputEvents() {
            List<ObjectElement> eventObjects = findOutputEventObjects(function.function.object)
                    .unique(Comparator.comparing { ObjectElement oE -> oE.getObjectDefinitionId() })
                    .sort { ObjectElement oE1, ObjectElement oE2 -> ModelUtils.getElementsCoordinatesComparator().compare(oE1, oE2) }
            outputEvents = eventObjects.collect { ObjectElement eventObject -> new CommonObjectInfo(eventObject) }
        }

        void findPerformers() {
            List<Edge> performerEdges = function.function.object.getEnterEdges()
                    .findAll { Edge e -> e.getEdgeTypeId() in PERFORMER_W_EPC_FUNCTION_EDGE_TYPE_IDS }

            for (edge in performerEdges) {
                ObjectElement objectElement = edge.getSource() as ObjectElement
                if (objectElement.getObjectDefinition().getObjectTypeId() in [BUSINESS_ROLE_OBJECT_TYPE_ID, GROUP_OBJECT_TYPE_ID, ORGANIZATIONAL_UNIT_OBJECT_TYPE_ID]) {
                    performers.add(new PerformerInfo(objectElement, edge))
                }
            }
        }

        void findInformationSystems() {
            List<ObjectElement> informationSystemObjects = function.function.object.getEnterEdges()
                    .findAll { Edge e -> e.getEdgeTypeId() in INFORMATION_SYSTEM_W_EPC_FUNCTION_EDGE_TYPE_IDS }
                    .collect { Edge e -> e.getSource() as ObjectElement }
                    .findAll { ObjectElement oE -> oE.getObjectDefinition().getObjectTypeId() == APPLICATION_SYSTEM_TYPE_OBJECT_TYPE_ID }
                    .unique(Comparator.comparing { ObjectElement oE -> oE.getObjectDefinitionId() })
                    .sort { ObjectElement oE1, ObjectElement oE2 -> ModelUtils.getElementsCoordinatesComparator().compare(oE1, oE2) }
            informationSystems = informationSystemObjects.collect { ObjectElement informationSystemObject -> new CommonObjectInfo(informationSystemObject, true) }
        }

        void findChildFunctions(List<EPCFunctionDescription> epcFunctions) {
            List<ObjectElement> eventObjects = findOutputEventObjects(function.function.object)
                    .sort { ObjectElement oE1, ObjectElement oE2 -> ModelUtils.getElementsCoordinatesComparator().compare(oE1, oE2) }

            List<ObjectElement> childFunctionObjects = []
            eventObjects.each { ObjectElement eventObject ->
                childFunctionObjects.addAll(
                        findChildFunctionsForEvent(eventObject)
                )
            }

            childFunctionObjects = childFunctionObjects
                    .unique(Comparator.comparing { ObjectElement oE -> oE.getObjectDefinitionId() })
                    .sort { ObjectElement oE1, ObjectElement oE2 -> ModelUtils.getElementsCoordinatesComparator().compare(oE1, oE2) }

            childFunctionObjects.each { ObjectElement childFunctionObject ->
                if (childFunctionObject.getSymbolId() in EXCLUDED_FUNCTION_SYMBOL_IDS) {
                    childExternalFunctions.add(new CommonFunctionInfo(childFunctionObject))
                } else {
                    EPCFunctionDescription childEPCFunction = epcFunctions
                            .find { EPCFunctionDescription epcFunction -> epcFunction.function.function.object.getId() == childFunctionObject.getId() }
                    childEPCFunctions.add(childEPCFunction)
                }
            }
        }

        private List<ObjectElement> findOutputEventObjects(ObjectElement functionObject) {
            List<ObjectElement> eventObjects = functionObject.getExitEdges()
                    .findAll { Edge e -> e.getEdgeTypeId() in EPC_FUNCTION_W_EVENT_EDGE_TYPE_IDS }
                    .collect { Edge e -> e.getTarget() as ObjectElement }
                    .findAll { ObjectElement oE -> oE.getObjectDefinition().getObjectTypeId() == EVENT_OBJECT_TYPE_ID }
                    .unique(Comparator.comparing { ObjectElement oE -> oE.getId() })

            List<ObjectElement> operators = functionObject.getExitEdges()
                    .findAll { Edge e -> e.getEdgeTypeId() in EPC_FUNCTION_W_OPERATOR_EDGE_TYPE_IDS }
                    .collect { Edge e -> e.getTarget() as ObjectElement }
                    .findAll { ObjectElement oE -> oE.getObjectDefinition().getObjectTypeId() == RULE_OBJECT_TYPE_ID }
                    .unique(Comparator.comparing { ObjectElement oE -> oE.getId() })

            while (operators) {
                List<ObjectElement> unparsedOperators = operators
                operators = []

                unparsedOperators.each { ObjectElement unparsedOperator ->
                    eventObjects.addAll(findOutputEventsViaOperator(unparsedOperator))
                    operators.addAll(findOutputOperatorsViaOperator(unparsedOperator))
                }
            }

            return eventObjects
        }

        private List<ObjectElement> findOutputEventsViaOperator(ObjectElement operator) {
            List<ObjectElement> eventObjects = operator.getExitEdges()
                    .findAll { Edge e -> e.getEdgeTypeId() in OPERATOR_W_EVENT_EDGE_TYPE_IDS }
                    .collect { Edge e -> e.getTarget() as ObjectElement }
                    .findAll { ObjectElement oE -> oE.getObjectDefinition().getObjectTypeId() == EVENT_OBJECT_TYPE_ID }
                    .unique(Comparator.comparing { ObjectElement oE -> oE.getId() })
            return eventObjects
        }

        private List<ObjectElement> findChildFunctionsForEvent(ObjectElement eventObject) {
            List<ObjectElement> childFunctionObjectsForEvent = eventObject.getExitEdges()
                    .findAll { Edge e -> e.getEdgeTypeId() == EVENT_W_EPC_FUNCTION_EDGE_TYPE_ID }
                    .collect { Edge e -> e.getTarget() as ObjectElement }
                    .findAll { ObjectElement oE -> oE.getObjectDefinition().getObjectTypeId() == FUNCTION_OBJECT_TYPE_ID }
                    .unique(Comparator.comparing { ObjectElement oE -> oE.getId() })
                    .sort { ObjectElement oE1, ObjectElement oE2 -> ModelUtils.getElementsCoordinatesComparator().compare(oE1, oE2) }

            List<ObjectElement> operators = eventObject.getExitEdges()
                    .findAll { Edge e -> e.getEdgeTypeId() == EVENT_W_OPERATOR_EDGE_TYPE_ID }
                    .collect { Edge e -> e.getTarget() as ObjectElement }
                    .findAll { ObjectElement oE -> oE.getObjectDefinition().getObjectTypeId() == RULE_OBJECT_TYPE_ID }
                    .unique(Comparator.comparing { ObjectElement oE -> oE.getId() })
                    .sort { ObjectElement oE1, ObjectElement oE2 -> ModelUtils.getElementsCoordinatesComparator().compare(oE1, oE2) }

            while (operators) {
                List<ObjectElement> unparsedOperators = operators
                operators = []

                unparsedOperators.each { ObjectElement unparsedOperator ->
                    childFunctionObjectsForEvent.addAll(findChildFunctionsViaOperator(unparsedOperator))
                    operators.addAll(findOutputOperatorsViaOperator(unparsedOperator))
                }
            }
            return childFunctionObjectsForEvent
        }

        private List<ObjectElement> findChildFunctionsViaOperator(ObjectElement operator) {
            List<ObjectElement> eventObjects = operator.getExitEdges()
                    .findAll { Edge e -> e.getEdgeTypeId() == OPERATOR_W_EPC_FUNCTION_EDGE_TYPE_ID }
                    .collect { Edge e -> e.getTarget() as ObjectElement }
                    .findAll { ObjectElement oE -> oE.getObjectDefinition().getObjectTypeId() == FUNCTION_OBJECT_TYPE_ID }
                    .unique(Comparator.comparing { ObjectElement oE -> oE.getId() })
            return eventObjects
        }

        private List<ObjectElement> findOutputOperatorsViaOperator(ObjectElement operator) {
            List<ObjectElement> operators = operator.getExitEdges()
                    .findAll { Edge e -> e.getEdgeTypeId() in OPERATOR_W_OPERATOR_EDGE_TYPE_IDS }
                    .collect { Edge e -> e.getTarget() as ObjectElement }
                    .findAll { ObjectElement oE -> oE.getObjectDefinition().getObjectTypeId() == RULE_OBJECT_TYPE_ID }
                    .unique(Comparator.comparing { ObjectElement oE -> oE.getId() })
            return operators
        }
    }

    private class InputFlowDescription {
        CommonObjectInfo inputFlow
        List<CommonFunctionInfo> suppliers = []

        InputFlowDescription(ObjectElement inputFlow, List<ObjectElement> supplierObjects) {
            this.inputFlow = new CommonObjectInfo(inputFlow)
            this.suppliers = supplierObjects.collect { ObjectElement supplierObject -> new CommonFunctionInfo(supplierObject) }
        }
    }

    private class OutputFlowDescription {
        CommonObjectInfo outputFlow
        public List<CommonFunctionInfo> customers = []

        OutputFlowDescription(ObjectElement outputFlow, List<ObjectElement> customerObjects) {
            this.outputFlow = new CommonObjectInfo(outputFlow)
            this.customers = customerObjects.collect { ObjectElement customerObject -> new CommonFunctionInfo(customerObject) }
        }
    }

    private class BusinessProcessRegulationDocument {
        String fileName
        SubprocessDescription subprocessDescription
        XWPFDocument document
        int detailLevel
        XWPFNumbering numbering

        FileNodeDTO content = null

        BusinessProcessRegulationDocument(String fileName, SubprocessDescription subprocessDescription, XWPFDocument template, int detailLevel) {
            this.fileName = fileName
            this.subprocessDescription = subprocessDescription
            this.document = template
            this.detailLevel = detailLevel
            this.numbering = document.numbering
        }

        void fillSimpleTexts() {
            Map<String, String> simpleTemplateMap = getSimpleTemplateMap()
            for (templateKey in simpleTemplateMap.keySet()) {
                String pattern = "<${templateKey}>"
                String replacement = simpleTemplateMap.get(templateKey)

                if (!replacement) {
                    continue
                }

                replaceParagraphsText(pattern, replacement)
                replaceHeadersText(pattern, replacement)
            }
        }

        private Map<String, String> getSimpleTemplateMap() {
            String docVersionWithDateTemplateValue
            if (docVersion && docDate) {
                docVersionWithDateTemplateValue = "${docVersion} от ${docDate}"
            } else if (docVersion && !docDate) {
                docVersionWithDateTemplateValue = "${docVersion} от <${DOC_DATE_TEMPLATE_KEY}>"
            } else if (!docVersion && docDate) {
                docVersionWithDateTemplateValue = "<${DOC_VERSION_TEMPLATE_KEY}> от ${docDate}"
            } else {
                docVersionWithDateTemplateValue = "<${DOC_VERSION_TEMPLATE_KEY}> от <${DOC_DATE_TEMPLATE_KEY}>"
            }

            Map<String, String> map = new HashMap<>()
            map.put(PROCESS_NAME_UPPER_CASE_TEMPLATE_KEY, subprocessDescription.subprocess.function.name.toUpperCase())
            map.put(PROCESS_CODE_TEMPLATE_KEY, subprocessDescription.subprocess.code)
            map.put(DOC_VERSION_WITH_DATE_TEMPLATE_KEY, docVersionWithDateTemplateValue)
            map.put(DOC_YEAR_TEMPLATE_KEY, currentYear)
            map.put(DOC_VERSION_TEMPLATE_KEY, docVersion)
            map.put(DOC_DATE_TEMPLATE_KEY, docDate)
            map.put(PROCESS_NAME_TEMPLATE_KEY, subprocessDescription.subprocess.function.name)
            map.put(FIRST_LEVEL_PROCESS_CODE_TEMPLATE_KEY, subprocessDescription.parentProcess.code)
            map.put(FIRST_LEVEL_PROCESS_NAME_TEMPLATE_KEY, subprocessDescription.parentProcess.function.name)
            map.put(PROCESS_REQUIREMENTS_TEMPLATE_KEY, subprocessDescription.subprocess.requirements ? "${subprocessDescription.subprocess.requirements}." : subprocessDescription.subprocess.requirements)
            return map
        }

        void fillLists() {
            fillProcessOwnersWithPosition()
            fillRequisitesNormativeDocuments()
            fillProcessOwners()
            fillProcessGoals()
        }

        private void fillProcessOwnersWithPosition() {
            String pattern = "<${PROCESS_OWNER_POSITION_TEMPLATE_KEY}> <${PROCESS_OWNER_TEMPLATE_KEY}>"
            String placeholderCopy = ", ${pattern}"
            for (owner in subprocessDescription.owners) {
                String replacement = ''

                if (owner.type == SubprocessOwnerType.ORGANIZATIONAL_UNIT) {
                    replacement += owner.leadershipPosition ? owner.leadershipPosition : "<${PROCESS_OWNER_POSITION_TEMPLATE_KEY}>"
                }

                if (owner.type == SubprocessOwnerType.GROUP) {
                    replacement += 'группа'
                }

                replacement += ' '
                replacement += owner.owner.name ? "(${owner.owner.name})" : "<${PROCESS_OWNER_TEMPLATE_KEY}>"

                if (pattern == replacement) {
                    continue
                }

                replacement += placeholderCopy
                replaceParagraphsText(pattern, replacement)
            }

            replaceParagraphsText(placeholderCopy, '')
        }

        private void fillRequisitesNormativeDocuments() {
            String pattern = "<${REQUISITES_NORMATIVE_DOCUMENT_TEMPLATE_KEY}>"
            List<String> requisitesNormativeDocuments = subprocessDescription.completedNormativeDocuments.collect { NormativeDocumentInfo normativeDocument -> normativeDocument.requisites ? normativeDocument.requisites : pattern }
            requisitesNormativeDocuments = requisitesNormativeDocuments.sort()

            int normativeDocumentsCount = requisitesNormativeDocuments.size()
            requisitesNormativeDocuments.eachWithIndex { String requisitesNormativeDocument, int number ->
                String replacement = number + 1 < normativeDocumentsCount ? "${requisitesNormativeDocument};" : "${requisitesNormativeDocument}."
                replaceInCopyParagraph(document, pattern, replacement)
            }

            if (listHasNotPatternValue(requisitesNormativeDocuments, pattern)) {
                Pattern searchPattern = Pattern.compile("^${pattern}\$")
                List<XWPFParagraph> paragraphs = findParagraphsByPattern(document, searchPattern)
                paragraphs.each { XWPFParagraph paragraph ->
                    removeParagraph(document, paragraph)
                }
            }
        }

        private void fillProcessOwners() {
            String pattern = "<${PROCESS_OWNER_TEMPLATE_KEY}>"
            String placeholderCopy = ", ${pattern}"
            for (owner in subprocessDescription.owners) {
                String replacement = owner.owner.name ? owner.owner.name : "<${PROCESS_OWNER_TEMPLATE_KEY}>"

                if (pattern == replacement) {
                    continue
                }

                replacement += placeholderCopy
                replaceParagraphsText(pattern, replacement)
            }

            replaceParagraphsText(placeholderCopy, '')
        }

        private void fillProcessGoals() {
            String pattern = "<${PROCESS_GOAL_TEMPLATE_KEY}>"

            int goalsCount = subprocessDescription.goals.size()
            subprocessDescription.goals.eachWithIndex { CommonObjectInfo goal, int number ->
                String goalName = goal.name ? goal.name : pattern
                String replacement = number + 1 < goalsCount ? "${goalName};" : "${goalName}."
                replaceInCopyParagraph(document, pattern, replacement)
            }

            if (subprocessDescription.goals) {
                Pattern searchPattern = Pattern.compile("^${pattern}\$")
                List<XWPFParagraph> paragraphs = findParagraphsByPattern(document, searchPattern)
                paragraphs.each { XWPFParagraph paragraph ->
                    removeParagraph(document, paragraph)
                }
            }
        }

        void fillTables() {
            fillAbbreviations()
            fillBusinessProcessHierarchy()
            fillExternalBusinessProcesses()
            fillProcessBusinessRoles()
            fillDocumentCollections()
        }

        private void fillAbbreviations() {
            XWPFTable table = findTableByHeaders(document, ABBREVIATION_TABLE_HEADERS)

            if (table.getRows().size() != 2) {
                return
            }

            String namePattern = "<${ABBREVIATION_TEMPLATE_KEY}>"
            String valuePattern = "<${ABBREVIATION_VALUE_TEMPLATE_KEY}>"
            for (abbreviation in foundedAbbreviations) {
                String nameReplacement = "${abbreviation.key}"
                String valueReplacement = "${abbreviation.value}"

                XWPFTableRow newTableRow = copyTableRow(table.getRows().get(1), table)

                replaceParagraphText(newTableRow.getTableCells().get(0).getParagraphs().get(0), namePattern, nameReplacement)
                replaceParagraphText(newTableRow.getTableCells().get(1).getParagraphs().get(0), valuePattern, valueReplacement)
            }

            table.removeRow(1)
        }

        private void fillBusinessProcessHierarchy() {
            XWPFTable table = findTableByHeaders(document, BUSINESS_PROCESS_HIERARCHY_TABLE_HEADERS)

            if (table.getRows().size() != 2) {
                return
            }

            fillBusinessProcess(table, '1', subprocessDescription.parentProcess.code, subprocessDescription.parentProcess.function.name)
            fillBusinessProcess(table, '2', subprocessDescription.subprocess.code, subprocessDescription.subprocess.function.name)

            for (scenarioDescription in subprocessDescription.scenarios) {
                fillBusinessProcess(table, '3', scenarioDescription.scenario.functionInfo.code, scenarioDescription.scenario.functionInfo.function.name)

                if (detailLevel == 3) {
                    continue
                }

                for (procedureDescription in scenarioDescription.procedures) {
                    fillBusinessProcess(table, '4', procedureDescription.procedure.functionInfo.code, procedureDescription.procedure.functionInfo.function.name)
                }
            }

            table.removeRow(1)
        }

        private void fillBusinessProcess(XWPFTable table, String level, String code, String name) {
            String levelPattern = "<${BUSINESS_PROCESS_LEVEL_TEMPLATE_KEY}>"
            String namePattern = "<${BUSINESS_PROCESS_CODE_TEMPLATE_KEY}> <${BUSINESS_PROCESS_NAME_TEMPLATE_KEY}>"

            code = code ? code : "<${BUSINESS_PROCESS_CODE_TEMPLATE_KEY}>"
            name = name ? name : "<${BUSINESS_PROCESS_NAME_TEMPLATE_KEY}>"
            String nameReplacement = "${code} ${name}"

            XWPFTableRow newTableRow = copyTableRow(table.getRows().get(1), table)
            replaceParagraphText(newTableRow.getTableCells().get(0).getParagraphs().get(0), levelPattern, level)
            replaceParagraphText(newTableRow.getTableCells().get(1).getParagraphs().get(0), namePattern, nameReplacement)
        }

        private void fillExternalBusinessProcesses() {
            XWPFTable table = findTableByHeaders(document, EXTERNAL_BUSINESS_PROCESS_TABLE_HEADERS)

            if (table.getRows().size() != 3) {
                return
            }

            List<SubprocessDescription.ExternalProcessDescription> sortedExternalProcessDescriptionsWithInputFlows = subprocessDescription.completedExternalProcessesWithInputFlows.sort { SubprocessDescription.ExternalProcessDescription ePD -> ePD.externalProcess.function.name }
            List<SubprocessDescription.ExternalProcessDescription> sortedExternalProcessDescriptionsWithOutputFlows = subprocessDescription.completedExternalProcessesWithOutputFlows.sort { SubprocessDescription.ExternalProcessDescription ePD -> ePD.externalProcess.function.name }

            for (externalProcessDescription in sortedExternalProcessDescriptionsWithInputFlows) {
                fillExternalBusinessProcess(table, externalProcessDescription, EXTERNAL_BUSINESS_PROCESS_INPUT_TEMPLATE_KEY, 1)
            }

            for (externalProcessDescription in sortedExternalProcessDescriptionsWithOutputFlows) {
                fillExternalBusinessProcess(table, externalProcessDescription, EXTERNAL_BUSINESS_PROCESS_OUTPUT_TEMPLATE_KEY, 2)
            }

            table.removeRow(2)
            table.removeRow(1)
        }

        private void fillExternalBusinessProcess(XWPFTable table, SubprocessDescription.ExternalProcessDescription externalProcessDescription, String flowTemplateKey, int flowColumnNumber) {
            String namePattern = "<${EXTERNAL_BUSINESS_PROCESS_CODE_TEMPLATE_KEY}> <${EXTERNAL_BUSINESS_PROCESS_NAME_TEMPLATE_KEY}>"
            String flowPattern = "<${flowTemplateKey}>"

            String code = externalProcessDescription.externalProcess.code ? externalProcessDescription.externalProcess.code : "<${EXTERNAL_BUSINESS_PROCESS_CODE_TEMPLATE_KEY}>"
            String name = externalProcessDescription.externalProcess.function.name ? externalProcessDescription.externalProcess.function.name : "<${EXTERNAL_BUSINESS_PROCESS_NAME_TEMPLATE_KEY}>"
            String nameReplacement = "${code} ${name}"

            XWPFTableRow newTableRow = copyTableRow(table.getRows().get(flowColumnNumber), table)
            replaceParagraphText(newTableRow.getTableCells().get(0).getParagraphs().get(0), namePattern, nameReplacement)

            List<String> flowNames = externalProcessDescription.flows.collect { CommonObjectInfo flow -> flow.name ? flow.name : "<${flowTemplateKey}>" }
            flowNames = flowNames.sort()

            int flowsCount = flowNames.size()
            flowNames.eachWithIndex { String flowName, int number ->
                String flowReplacement = number + 1 < flowsCount ? "${flowName};" : flowName
                replaceInCopyParagraph(newTableRow.getTableCells().get(flowColumnNumber), flowPattern, flowReplacement)
            }

            newTableRow.getTableCells().get(flowColumnNumber).removeParagraph(newTableRow.getTableCells().get(flowColumnNumber).getParagraphs().size() - 1)
        }

        private void fillProcessBusinessRoles() {
            if (detailLevel == 3) {
                List<IBodyElement> elementsToDelete = findBodyElements(document, PROCESS_BUSINESS_ROLES_PARAGRAPH_TEXT_TO_FIND, PROCESS_PARTICIPANTS_PARAGRAPH_TEXT_TO_FIND)
                removeBodyElements(document, elementsToDelete)
            }

            if (detailLevel == 4) {
                List<XWPFParagraph> paragraphsToDelete = findParagraphsByText(document, PROCESS_PARTICIPANTS_PARAGRAPH_TEXT_TO_FIND)

                if (paragraphsToDelete.size() != 1) {
                    throw new Exception('Неверное количество параграфов для удаления в пункте "Участники процеса"')
                }

                int posOfParagraph = document.getPosOfParagraph(paragraphsToDelete[0])
                document.removeBodyElement(posOfParagraph)

                XWPFTable table = findTableByHeaders(document, PROCESS_BUSINESS_ROLE_TABLE_HEADERS)

                if (table.getRows().size() != 2) {
                    return
                }

                List<BusinessRoleInfo> sortedBusinessRoles = subprocessDescription.completedBusinessRoles.sort { BusinessRoleInfo bRI -> bRI.businessRole.name }
                for (businessRoleInfo in sortedBusinessRoles) {
                    String businessRolePattern = "<${PROCESS_BUSINESS_ROLE_TEMPLATE_KEY}>"
                    String businessRolePositionPattern = "<${PROCESS_BUSINESS_ROLE_POSITION_TEMPLATE_KEY}>"
                    String businessRolePositionOrganizationalUnitPattern = "<${PROCESS_BUSINESS_ROLE_POSITION_ORGANIZATIONAL_UNIT_TEMPLATE_KEY}>"
                    String positionPattern = "${businessRolePositionPattern} (${businessRolePositionOrganizationalUnitPattern})"

                    String businessRoleReplacement = businessRoleInfo.businessRole.name ? businessRoleInfo.businessRole.name : "<${PROCESS_BUSINESS_ROLE_TEMPLATE_KEY}>"

                    XWPFTableRow newTableRow = copyTableRow(table.getRows().get(1), table)
                    replaceParagraphText(newTableRow.getTableCells().get(0).getParagraphs().get(0), businessRolePattern, businessRoleReplacement)

                    List<String> positions = []
                    for (position in businessRoleInfo.positions) {
                        String positionName = position.position.name ? position.position.name : businessRolePositionPattern

                        String organizationalUnitName
                        if (position.organizationalUnit) {
                            organizationalUnitName = position.organizationalUnit.name ? position.organizationalUnit.name : businessRolePositionOrganizationalUnitPattern
                        } else {
                            organizationalUnitName = businessRolePositionOrganizationalUnitPattern
                        }

                        positions.add("${positionName} (${organizationalUnitName})")
                    }
                    positions = positions.sort()

                    int positionsCount = positions.size()
                    positions.eachWithIndex { String position, int number ->
                        String positionReplacement = number + 1 < positionsCount ? "${position};" : position
                        replaceInCopyParagraph(newTableRow.getTableCells().get(1), positionPattern, positionReplacement)
                    }

                    if (positions) {
                        newTableRow.getTableCells().get(1).removeParagraph(newTableRow.getTableCells().get(1).getParagraphs().size() - 1)
                    } else {
                        addParagraphText(newTableRow.getTableCells().get(1).getParagraphArray(0), '')
                    }
                }

                table.removeRow(1)
            }
        }

        private void fillDocumentCollections() {
            XWPFTable table = findTableByHeaders(document, DOCUMENT_COLLECTION_TABLE_HEADERS)

            if (table.getRows().size() != 2) {
                return
            }

            for (documentCollectionInfo in subprocessDescription.completedDocumentCollections) {
                String documentCollectionPattern = "<${PROCESS_DOCUMENT_COLLECTION_TEMPLATE_KEY}>"
                String containedDocumentPattern = "<${PROCESS_DOCUMENT_COLLECTION_CONTAINED_DOCUMENT_TEMPLATE_KEY}>"

                String documentCollectionReplacement = documentCollectionInfo.collection.document.name ? documentCollectionInfo.collection.document.name : "<${PROCESS_DOCUMENT_COLLECTION_TEMPLATE_KEY}>"
                documentCollectionReplacement += " [${documentCollectionInfo.collection.type}]"

                XWPFTableRow newTableRow = copyTableRow(table.getRows().get(1), table)
                replaceParagraphText(newTableRow.getTableCells().get(0).getParagraphs().get(0), documentCollectionPattern, documentCollectionReplacement)

                List<String> containedDocuments = documentCollectionInfo.containedDocuments.collect { DocumentInfo containedDocument -> (containedDocument.document.name ? containedDocument.document.name : containedDocumentPattern) + " [${containedDocument.type}]" }
                containedDocuments = containedDocuments.sort()

                int containedDocumentsCount = containedDocuments.size()
                containedDocuments.eachWithIndex { String containedDocument, int number ->
                    String containedDocumentReplacement = number + 1 < containedDocumentsCount ? "${containedDocument};" : containedDocument
                    replaceInCopyParagraph(newTableRow.getTableCells().get(1), containedDocumentPattern, containedDocumentReplacement)
                }

                if (containedDocuments) {
                    newTableRow.getTableCells().get(1).removeParagraph(newTableRow.getTableCells().get(1).getParagraphs().size() - 1)
                } else {
                    addParagraphText(newTableRow.getTableCells().get(1).getParagraphArray(0), '')
                }
            }

            table.removeRow(1)
        }

        void fillModels() {
            fillProcessModel()
            fillScenarioModels()
        }

        private void fillProcessModel() {
            String processModelPattern = "<${PROCESS_MODEL_TEMPLATE_KEY}>"
            List<XWPFParagraph> paragraphs = findParagraphsByText(document, processModelPattern)

            if (paragraphs.size() != 1) {
                throw new Exception('Неверное количество параграфов модели процесса')
            }

            XWPFParagraph imageParagraph = paragraphs[0]
            // позиция относительно всех элемнтов
            int imageParagraphPosition = document.getPosOfParagraph(imageParagraph)
            // позиция относительно параграфов
            int imageParagraphSpecificPos = document.getParagraphPos(imageParagraphPosition)
            // удаление дополнительного разрыва раздела, так как:
            // 1 - при вставке изображения, после страницы изображения разрыв раздела генерируется автоматически
            // 2 - при удалении шаблона изображения, разрыв больше не нужен
            document.removeBodyElement(imageParagraphPosition + 2)

            if (subprocessDescription.processSelectionModel) {
                while (imageParagraph.getRuns().size() > 0) {
                    imageParagraph.removeRun(0)
                }

                XWPFParagraph labelParagraph = document.getParagraphArray(imageParagraphSpecificPos + 1)
                addPicture(imageParagraph, subprocessDescription.processSelectionModel, labelParagraph)
            } else {
                document.removeBodyElement(imageParagraphPosition + 1)
                document.removeBodyElement(imageParagraphPosition)
                document.removeBodyElement(imageParagraphPosition - 1)
                addParagraphText(document.getParagraphArray(imageParagraphSpecificPos - 2), '')
            }
        }

        private void fillScenarioModels() {
            if (detailLevel == 3) {
                List<IBodyElement> elementsToDelete = findBodyElements(document, RESPONSIBILITY_MATRIX_PARAGRAPH_TEXT_TO_FIND, DOCUMENT_COLLECTION_PARAGRAPH_TEXT_TO_FIND)
                removeBodyElements(document, elementsToDelete)
            } else if (detailLevel == 4) {
                List<IBodyElement> elementsToDelete = findBodyElements(document, SCENARIO_FUNCTIONS_TABLE_PARAGRAPH_TEXT_TO_FIND, RESPONSIBILITY_MATRIX_PARAGRAPH_TEXT_TO_FIND)
                removeBodyElements(document, elementsToDelete)
            }

            int scenarioNumber = 1
            int procedureNumber = 1
            for (scenarioDescription in subprocessDescription.scenarios) {
                List<IBodyElement> scenarioElements = findBodyElements(document, SCENARIO_PARAGRAPH_TEXT_TO_FIND, DOCUMENT_COLLECTION_PARAGRAPH_TEXT_TO_FIND, 1)
                List<IBodyElement> newScenarioElements = copyIBodyElements(scenarioElements)
                setNewNumbering(newScenarioElements)

                fillScenario(scenarioDescription, scenarioNumber, scenarioElements)

                if (detailLevel == 4) {
                    for (procedureDescription in scenarioDescription.procedures) {
                        List<IBodyElement> procedureElements = findBodyElements(document, PROCEDURE_PARAGRAPH_TEXT_TO_FIND, SCENARIO_PARAGRAPH_TEXT_TO_FIND, 1)
                        List<IBodyElement> newProcedureElements = copyIBodyElements(procedureElements)
                        setNewNumbering(newProcedureElements)

                        fillProcedure(procedureDescription, procedureNumber, procedureElements)

                        procedureNumber += 1
                    }

                    List<IBodyElement> elementsToDelete = findBodyElements(document, PROCEDURE_PARAGRAPH_TEXT_TO_FIND, SCENARIO_PARAGRAPH_TEXT_TO_FIND, 1)
                    removeBodyElements(document, elementsToDelete)
                }

                scenarioNumber += 1
            }

            List<IBodyElement> elementsToDelete = findBodyElements(document, SCENARIO_PARAGRAPH_TEXT_TO_FIND, DOCUMENT_COLLECTION_PARAGRAPH_TEXT_TO_FIND, 1)
            removeBodyElements(document, elementsToDelete)
        }

        private void fillScenario(ScenarioDescription description, int number, List<IBodyElement> elements) {
            String scenarioPattern = "<${SCENARIO_CODE_TEMPLATE_KEY}> <${SCENARIO_NAME_TEMPLATE_KEY}>"
            String scenarioCode = description.scenario.functionInfo.code ? description.scenario.functionInfo.code : "<${SCENARIO_CODE_TEMPLATE_KEY}>"
            String scenarioName = description.scenario.functionInfo.function.name ? description.scenario.functionInfo.function.name : "<${SCENARIO_NAME_TEMPLATE_KEY}>"
            String scenarioReplacement = "${scenarioCode} ${scenarioName}"

            if (scenarioPattern == scenarioReplacement) {
                throw new Exception("Сценрий [${description.scenario.functionInfo.function.object.getObjectDefinitionId()}] должен иметь либо код, либо имя")
            }

            replaceParagraphsText(elements, scenarioPattern, scenarioReplacement)

            String requirementsPattern = "<${SCENARIO_REQUIREMENTS_TEMPLATE_KEY}>"
            String requirementsReplacement = description.scenario.functionInfo.requirements ? "${description.scenario.functionInfo.requirements}." : requirementsPattern
            replaceParagraphsText(elements, requirementsPattern, requirementsReplacement)

            String pictureNumberPattern = "<${SCENARIO_PICTURE_NUMBER_TEMPLATE_KEY}>"
            String pictureNumberReplacement = "<${SCENARIO_PICTURE_NUMBER_TEMPLATE_KEY} ${number}>"
            replaceParagraphsText(elements, pictureNumberPattern, pictureNumberReplacement)

            String functionsTableNumberPattern = "<${SCENARIO_FUNCTIONS_TABLE_NUMBER_TEMPLATE_KEY}>"
            String functionsTableReplacement = "<${SCENARIO_FUNCTIONS_TABLE_NUMBER_TEMPLATE_KEY} ${number}>"
            replaceParagraphsText(elements, functionsTableNumberPattern, functionsTableReplacement)

            String responsibilityMatrixNumberPattern = "<${SCENARIO_RESPONSIBILITY_MATRIX_TABLE_NUMBER_TEMPLATE_KEY}>"
            String responsibilityMatrixReplacement = "<${SCENARIO_RESPONSIBILITY_MATRIX_TABLE_NUMBER_TEMPLATE_KEY} ${number}>"
            replaceParagraphsText(elements, responsibilityMatrixNumberPattern, responsibilityMatrixReplacement)

            String modelPattern = "<${SCENARIO_MODEL_TEMPLATE_KEY}>"
            List<XWPFParagraph> paragraphs = findParagraphsByText(document, modelPattern)

            if (paragraphs.size() != 2) {
                throw new Exception('Неверное количество параграфов модели сценария')
            }

            XWPFParagraph imageParagraph = paragraphs[0]
            // позиция относительно всех элемнтов
            int imageParagraphPosition = document.getPosOfParagraph(imageParagraph)
            // позиция относительно параграфов
            int imageParagraphSpecificPos = document.getParagraphPos(imageParagraphPosition)
            // удаление дополнительного разрыва раздела, так как:
            // 1 - при вставке изображения, после страницы изображения разрыв раздела генерируется автоматически
            // 2 - при удалении шаблона изображения, разрыв больше не нужен
            document.removeBodyElement(imageParagraphPosition + 2)

            while (imageParagraph.getRuns().size() > 0) {
                imageParagraph.removeRun(0)
            }

            XWPFParagraph labelParagraph = document.getParagraphArray(imageParagraphSpecificPos + 1)
            addPicture(imageParagraph, description.scenario.model, labelParagraph)

            if (detailLevel == 3) {
                XWPFTable table = findTableByHeaders(elements, FUNCTIONS_TABLE_HEADERS)
                List<EPCFunctionDescription> functions = description.scenario.epcFunctions

                if (table.getRows().size() != 8) {
                    return
                }

                fillFunctionsTable(table, functions)

                for (int pos = 7; pos > 0; pos--) {
                    table.removeRow(pos)
                }

                if (functions) {
                    table.removeRow(0)
                } else {
                    replaceParagraphText(table.getRows().get(0).getTableCells().get(0).getParagraphs().get(0), "<${FUNCTION_NUMBER_TEMPLATE_KEY}>", '')
                    replaceParagraphText(table.getRows().get(0).getTableCells().get(1).getParagraphs().get(0), "<${FUNCTION_CODE_TEMPLATE_KEY}> <${FUNCTION_NAME_TEMPLATE_KEY}>", '')
                }
            }

            if (detailLevel == 4) {
                XWPFTable table = findTableByHeaders(elements, RESPONSIBILITY_MATRIX_HEADERS)

                if (table.getRows().size() != 2) {
                    return
                }

                fillResponsibilityMatrix(table, description)

                table.removeRow(1)
            }
        }

        private static void fillResponsibilityMatrix(XWPFTable table, ScenarioDescription scenarioDescription) {
            XWPFTableRow headersRow = table.getRows().get(0)
            String businessRolePattern = "<${BUSINESS_ROLE_NAME_TEMPLATE_KEY}>"

            List<BusinessRoleInfo> businessRoles = scenarioDescription.getAllBusinessRoles()
            XWPFTableCell sourceBusinessRoleCell = headersRow.getTableCells().get(1)
            for (businessRole in businessRoles) {
                XWPFTableCell targetCell = headersRow.createCell()
                targetCell = copyTableCell(sourceBusinessRoleCell, targetCell)

                String businessRoleReplacement = businessRole.businessRole.name ? businessRole.businessRole.name : businessRolePattern
                replaceParagraphText(targetCell.getParagraphs().get(0), businessRolePattern, businessRoleReplacement)
            }

            if (businessRoles) {
                headersRow.removeCell(1)

                XWPFTableRow procedureRow = table.getRows().get(1)
                XWPFTableCell sourceValueCell = procedureRow.getTableCells().get(1)
                for (int cellNumber = 2; cellNumber < headersRow.getTableCells().size(); cellNumber++) {
                    XWPFTableCell targetCell = procedureRow.createCell()
                    copyTableCell(sourceValueCell, targetCell)
                }
            }

            for (procedure in scenarioDescription.procedures) {
                XWPFTableRow newProcedureRow = copyTableRow(table.getRows().get(1), table)

                String procedurePattern = "<${PROCEDURE_NAME_TEMPLATE_KEY}>"
                String procedureReplacement = procedure.procedure.functionInfo.function.name ? procedure.procedure.functionInfo.function.name : procedurePattern
                replaceParagraphText(newProcedureRow.getTableCells().get(0).getParagraphs().get(0), procedurePattern, procedureReplacement)

                List<String> procedureBusinessRoles = procedure.businessRoles.collect { BusinessRoleInfo businessRole -> (businessRole.businessRole.name ? businessRole.businessRole.name : businessRolePattern) }
                businessRoles.eachWithIndex { BusinessRoleInfo businessRole, int number ->
                    String businessRoleReplacement = businessRole.businessRole.name ? businessRole.businessRole.name : businessRolePattern

                    if (!(businessRoleReplacement in procedureBusinessRoles)) {
                        addParagraphText(newProcedureRow.getTableCells().get(number + 1).getParagraphs().get(0), '')
                    }
                }
            }
        }

        private void fillProcedure(ProcedureDescription description, int number, List<IBodyElement> elements) {
            String procedurePattern = "<${PROCEDURE_CODE_TEMPLATE_KEY}> <${PROCEDURE_NAME_TEMPLATE_KEY}>"
            String procedureCode = description.procedure.functionInfo.code ? description.procedure.functionInfo.code : "<${PROCEDURE_CODE_TEMPLATE_KEY}>"
            String procedureName = description.procedure.functionInfo.function.name ? description.procedure.functionInfo.function.name : "<${PROCEDURE_NAME_TEMPLATE_KEY}>"
            String procedureReplacement = "${procedureCode} ${procedureName}"

            if (procedurePattern == procedureReplacement) {
                throw new Exception("Процедура [${description.procedure.functionInfo.function.object.getObjectDefinitionId()}] должна иметь либо код, либо имя")
            }

            replaceParagraphsText(elements, procedurePattern, procedureReplacement)

            String requirementsPattern = "<${PROCEDURE_REQUIREMENTS_TEMPLATE_KEY}>"
            String requirementsReplacement = description.procedure.functionInfo.requirements ? "${description.procedure.functionInfo.requirements}." : requirementsPattern
            replaceParagraphsText(elements, requirementsPattern, requirementsReplacement)

            String pictureNumberPattern = "<${PROCEDURE_PICTURE_NUMBER_TEMPLATE_KEY}>"
            String pictureNumberReplacement = "<${PROCEDURE_PICTURE_NUMBER_TEMPLATE_KEY} ${number}>"
            replaceParagraphsText(elements, pictureNumberPattern, pictureNumberReplacement)

            String functionsTableNumberPattern = "<${PROCEDURE_FUNCTIONS_TABLE_NUMBER_TEMPLATE_KEY}>"
            String functionsTableReplacement = "<${PROCEDURE_FUNCTIONS_TABLE_NUMBER_TEMPLATE_KEY} ${number}>"
            replaceParagraphsText(elements, functionsTableNumberPattern, functionsTableReplacement)

            String modelPattern = "<${PROCEDURE_MODEL_TEMPLATE_KEY}>"
            List<XWPFParagraph> paragraphs = findParagraphsByText(document, modelPattern)

            if (paragraphs.size() != 3) {
                throw new Exception('Неверное количество параграфов модели процедуры')
            }

            XWPFParagraph imageParagraph = paragraphs[0]
            // позиция относительно всех элемнтов
            int imageParagraphPosition = document.getPosOfParagraph(imageParagraph)
            // позиция относительно параграфов
            int imageParagraphSpecificPos = document.getParagraphPos(imageParagraphPosition)
            // удаление дополнительного разрыва раздела, так как:
            // 1 - при вставке изображения, после страницы изображения разрыв раздела генерируется автоматически
            // 2 - при удалении шаблона изображения, разрыв больше не нужен
            document.removeBodyElement(imageParagraphPosition + 2)

            while (imageParagraph.getRuns().size() > 0) {
                imageParagraph.removeRun(0)
            }

            XWPFParagraph labelParagraph = document.getParagraphArray(imageParagraphSpecificPos + 1)
            addPicture(imageParagraph, description.procedure.model, labelParagraph)

            XWPFTable table = findTableByHeaders(elements, FUNCTIONS_TABLE_HEADERS)
            List<EPCFunctionDescription> functions = description.procedure.epcFunctions

            if (table.getRows().size() != 8) {
                return
            }

            fillFunctionsTable(table, functions)

            for (int pos = 7; pos > 0; pos--) {
                table.removeRow(pos)
            }

            if (functions) {
                table.removeRow(0)
            } else {
                replaceParagraphText(table.getRows().get(0).getTableCells().get(0).getParagraphs().get(0), "<${FUNCTION_NUMBER_TEMPLATE_KEY}>", '')
                replaceParagraphText(table.getRows().get(0).getTableCells().get(1).getParagraphs().get(0), "<${FUNCTION_CODE_TEMPLATE_KEY}> <${FUNCTION_NAME_TEMPLATE_KEY}>", '')
            }
        }

        private static void fillFunctionsTable(XWPFTable table, List<EPCFunctionDescription> functions) {
            for (function in functions) {
                XWPFTableRow functionRow = copyTableRow(table.getRows().get(0), table)
                XWPFTableRow performersRow = copyTableRow(table.getRows().get(1), table)
                XWPFTableRow inputsRow = copyTableRow(table.getRows().get(2), table)
                XWPFTableRow informationSystemsRow = copyTableRow(table.getRows().get(3), table)
                XWPFTableRow requirementsRow = copyTableRow(table.getRows().get(4), table)
                XWPFTableRow outputsRow = copyTableRow(table.getRows().get(5), table)
                XWPFTableRow durationRow = copyTableRow(table.getRows().get(6), table)
                XWPFTableRow childFunctionsRow = copyTableRow(table.getRows().get(7), table)

                fillFunctionNumber(function, functionRow.getTableCells().get(0))
                fillFunctionName(function, functionRow.getTableCells().get(1))
                fillPerformers(function, performersRow.getTableCells().get(1))
                fillFunctionInputs(function, inputsRow.getTableCells().get(1))
                fillInformationSystems(function, informationSystemsRow.getTableCells().get(1))
                fillRequirements(function, requirementsRow.getTableCells().get(1))
                fillFunctionOutputs(function, outputsRow.getTableCells().get(1))
                fillDuration(function, durationRow.getTableCells().get(1))
                fillChildFunctions(function, childFunctionsRow.getTableCells().get(1))
            }
        }

        private static void fillFunctionNumber(EPCFunctionDescription function, XWPFTableCell functionTableCell) {
            String numberPattern = "<${FUNCTION_NUMBER_TEMPLATE_KEY}>"
            String numberReplacement = function.number.toString()
            replaceParagraphText(functionTableCell.getParagraphs().get(0), numberPattern, numberReplacement)
        }

        private static void fillFunctionName(EPCFunctionDescription function, XWPFTableCell functionTableCell) {
            String functionPattern = "<${FUNCTION_CODE_TEMPLATE_KEY}> <${FUNCTION_NAME_TEMPLATE_KEY}>"
            String functionCode = function.function.code ? function.function.code : "<${FUNCTION_CODE_TEMPLATE_KEY}>"
            String functionName = function.function.function.name ? function.function.function.name : "<${FUNCTION_NAME_TEMPLATE_KEY}>"
            String functionReplacement = "${functionCode} ${functionName}"
            replaceParagraphText(functionTableCell.getParagraphs().get(0), functionPattern, functionReplacement)
        }

        private static void fillPerformers(EPCFunctionDescription function, XWPFTableCell functionTableCell) {
            String performerPattern = "<${PERFORMER_TEMPLATE_KEY}>"
            List<String> performers = []
            for (performerInfo in function.performers) {
                String performer = ''
                performer += performerInfo.performer.name ? performerInfo.performer.name : performerPattern
                performer += " [${performerInfo.action}]"
                performers.add(performer)
            }
            performers = performers.sort()
            fillFunctionTableCell(functionTableCell, performers, performerPattern)
        }

        private static void fillFunctionInputs(EPCFunctionDescription function, XWPFTableCell functionTableCell) {
            String inputPattern = "<${INPUT_DOCUMENT_TEMPLATE_KEY}/${INPUT_EVENT_TEMPLATE_KEY}>"
            List<String> inputs = getFunctionDocuments(function.inputDocuments, "<${INPUT_DOCUMENT_TEMPLATE_KEY}>")

            if (inputs.isEmpty()) {
                inputs = getFunctionEvents(function.inputEvents, "<${INPUT_EVENT_TEMPLATE_KEY}>")
            }

            fillFunctionTableCell(functionTableCell, inputs, inputPattern)
        }

        private static void fillInformationSystems(EPCFunctionDescription function, XWPFTableCell functionTableCell) {
            String informationSystemPattern = "<${INFORMATION_SYSTEM_TEMPLATE_KEY}>"
            List<String> informationSystems = []
            for (informationSystemInfo in function.informationSystems) {
                String informationSystem = ''
                informationSystem += informationSystemInfo.name ? informationSystemInfo.name : informationSystemPattern
                informationSystems.add(informationSystem)
            }
            informationSystems = informationSystems.sort()
            fillFunctionTableCell(functionTableCell, informationSystems, informationSystemPattern)
        }

        private static void fillRequirements(EPCFunctionDescription function, XWPFTableCell functionTableCell) {
            String requirementsPattern = "<${FUNCTION_REQUIREMENTS_TEMPLATE_KEY}>"
            String requirementsReplacement = function.function.requirements
            replaceParagraphText(functionTableCell.getParagraphs().get(0), requirementsPattern, requirementsReplacement)
        }

        private static void fillFunctionOutputs(EPCFunctionDescription function, XWPFTableCell functionTableCell) {
            String outputPattern = "<${OUTPUT_DOCUMENT_TEMPLATE_KEY}/${OUTPUT_EVENT_TEMPLATE_KEY}>"
            List<String> outputs = getFunctionDocuments(function.outputDocuments, "<${OUTPUT_DOCUMENT_TEMPLATE_KEY}>")

            if (outputs.isEmpty()) {
                outputs = getFunctionEvents(function.outputEvents, "<${OUTPUT_EVENT_TEMPLATE_KEY}>")
            }

            fillFunctionTableCell(functionTableCell, outputs, outputPattern)
        }

        private static List<String> getFunctionEvents(List<CommonObjectInfo> events, String eventPattern) {
            List<String> resultEvents = []
            for (event in events) {
                String resultEvent = event.name ? event.name : eventPattern
                resultEvent += ' [событие]'
                resultEvents.add(resultEvent)
            }
            return resultEvents
        }

        private static List<String> getFunctionDocuments(List<DocumentInfo> documents, String documentPattern) {
            List<String> resultDocuments = []
            for (document in documents) {
                String resultDocument = ''
                resultDocument += document.document.name ? document.document.name : documentPattern
                resultDocument += " [${document.type}]"
                resultDocument += document.statuses ? " (${String.join(', ', document.statuses.collect { CommonObjectInfo status -> status.name })})" : ''
                resultDocuments.add(resultDocument)
            }
            return resultDocuments
        }

        private static void fillDuration(EPCFunctionDescription function, XWPFTableCell functionTableCell) {
            String durationPattern = "<${DURATION_TEMPLATE_KEY}>"
            String durationReplacement = function.duration
            replaceParagraphText(functionTableCell.getParagraphs().get(0), durationPattern, durationReplacement)
        }

        private static void fillChildFunctions(EPCFunctionDescription function, XWPFTableCell functionTableCell) {
            String childFunctionPattern = "<${CHILD_FUNCTION_TEMPLATE_KEY}>"
            List<String> childFunctions = []
            for (childEPCFunction in function.childEPCFunctions) {
                String childFunction = "Переход к п. ${childEPCFunction.number.toString()}"
                childFunctions.add(childFunction)
            }

            for (childExternalFunction in function.childExternalFunctions) {
                String childFunction = "Переход к процессу «${childExternalFunction.function.name}»"
                childFunctions.add(childFunction)
            }
            fillFunctionTableCell(functionTableCell, childFunctions, childFunctionPattern)
        }

        private static void fillFunctionTableCell(XWPFTableCell functionTableCell, List<String> elements, String pattern) {
            int elementsCount = elements.size()
            elements.eachWithIndex { String element, int number ->
                String replacement = number + 1 < elementsCount ? "${element};" : element
                replaceInCopyParagraph(functionTableCell, pattern, replacement)
            }

            if (elements) {
                functionTableCell.removeParagraph(functionTableCell.getParagraphs().size() - 1)
            } else {
                addParagraphText(functionTableCell.getParagraphArray(0), '')
            }
        }

        private void setNewNumbering(List<IBodyElement> elements) {
            BigInteger numID = null
            for (element in elements) {
                if (element.getElementType() == BodyElementType.PARAGRAPH) {
                    XWPFParagraph paragraph = (XWPFParagraph) element
                    String text = paragraph.getText()

                    if (text.startsWith(REQUIREMENTS_PARAGRAPH_TEXT_TO_FIND)) {
                        numID = getNewSimpleNumberingId()
                    }

                    if (text.startsWith(REQUIREMENTS_PARAGRAPH_TEXT_TO_FIND)
                            || text.startsWith(PICTURE_PARAGRAPH_TEXT_TO_FIND)
                            || text.startsWith(FUNCTIONS_PARAGRAPH_TEXT_TO_FIND)
                            || text.startsWith(RESPONSIBILITY_MATRIX_PARAGRAPH_TEXT_TO_FIND)
                    ) {
                        paragraph.setNumID(numID)
                    }
                }
            }
        }

        private BigInteger getNewSimpleNumberingId() {
            CTAbstractNum cTAbstractNum = CTAbstractNum.Factory.newInstance()
            cTAbstractNum.setAbstractNumId(BigInteger.valueOf(0))

            CTLvl cTLvl = cTAbstractNum.addNewLvl()
            cTLvl.setIlvl(BigInteger.valueOf(0))
            cTLvl.addNewNumFmt().setVal(STNumberFormat.DECIMAL)
            cTLvl.addNewLvlText().setVal("%1.")
            cTLvl.addNewLvlJc().setVal(STJc.LEFT)
            cTLvl.addNewStart().setVal(BigInteger.valueOf(1))

            cTLvl.addNewRPr()
            CTFonts f = cTLvl.getRPr().addNewRFonts()
            f.setAscii('Times New Roman')
            f.setHAnsi('Times New Roman')

            XWPFAbstractNum abstractNum = new XWPFAbstractNum(cTAbstractNum)
            BigInteger abstractNumId = numbering.addAbstractNum(abstractNum)
            BigInteger numId = numbering.addNum(abstractNumId)

            XWPFNum num = numbering.getNum(numId)
            CTNumLvl lvlOverride = num.getCTNum().addNewLvlOverride()
            lvlOverride.setIlvl(BigInteger.ZERO)
            CTDecimalNumber number = lvlOverride.addNewStartOverride()
            number.setVal(BigInteger.ONE)

            return numId
        }

        void setLabelNumbers() {
            setTableNumbers()
            setPictureNumbers()
        }

        private void setTableNumbers() {
            Pattern tableNumberPattern = Pattern.compile('Таблица <(.*?)>')
            List<XWPFParagraph> tableLabelParagraphs = findParagraphsByPattern(document, tableNumberPattern)
            Map<String, String> numberTemplateBookmarkMap = setNumbers(document, tableLabelParagraphs, tableNumberPattern, 'Table')

            for (paragraph in tableLabelParagraphs) {
                removeParagraph(document, paragraph)
            }

            for (numberTemplate in numberTemplateBookmarkMap.keySet()) {
                List<XWPFParagraph> tableNumberParagraphs = findParagraphsByText(document, numberTemplate)
                setRefsToNumber(document, tableNumberParagraphs, numberTemplate, numberTemplateBookmarkMap.get(numberTemplate))

                for (paragraph in tableNumberParagraphs) {
                    removeParagraph(document, paragraph)
                }
            }
        }

        private void setPictureNumbers() {
            Pattern pictureNumberPattern = Pattern.compile('Рисунок <(.*?)>')
            List<XWPFParagraph> pictureLabelParagraphs = findParagraphsByPattern(document, pictureNumberPattern)
            Map<String, String> numberTemplateBookmarkMap = setNumbers(document, pictureLabelParagraphs, pictureNumberPattern, 'Picture')

            for (paragraph in pictureLabelParagraphs) {
                removeParagraph(document, paragraph)
            }

            for (numberTemplate in numberTemplateBookmarkMap.keySet()) {
                List<XWPFParagraph> pictureNumberParagraphs = findParagraphsByText(document, numberTemplate)
                setRefsToNumber(document, pictureNumberParagraphs, numberTemplate, numberTemplateBookmarkMap.get(numberTemplate))

                for (paragraph in pictureNumberParagraphs) {
                    removeParagraph(document, paragraph)
                }
            }
        }

        private static Map<String, String> setNumbers(IBody body, List<XWPFParagraph> paragraphs, Pattern templatePattern, String name) {
            Map<String, String> numberTemplateBookmarkMap = new HashMap<>()
            int number = 0
            for (sourceParagraph in paragraphs) {
                number += 1

                XmlCursor cursor = sourceParagraph.getCTP().newCursor()
                XWPFParagraph newParagraph = body.insertNewParagraph(cursor)
                CTPPr pPr = newParagraph.getCTP().isSetPPr() ? newParagraph.getCTP().getPPr() : newParagraph.getCTP().addNewPPr()
                pPr.set(sourceParagraph.getCTP().getPPr())

                boolean templateFound = false
                boolean templateParsed = false
                for (sourceRun in sourceParagraph.getRuns()) {
                    String sourceRunText = sourceRun.getText(0)

                    if (templateParsed) {
                        XWPFRun targetRun = newParagraph.createRun()
                        copyRun(sourceRun, targetRun)
                        continue
                    }

                    if (!templateFound && sourceRunText && sourceRunText.contains('<')) {
                        XWPFRun targetRun = newParagraph.createRun()
                        copyRun(sourceRun, targetRun)
                        String newText = sourceRunText.substring(0, sourceRunText.size() - 1)
                        targetRun.setText(newText, 0)
                        templateFound = true
                        continue
                    }

                    if (!templateParsed && sourceRunText && sourceRunText.contains('>')) {
                        XWPFRun targetRun = newParagraph.createRun()
                        copyRun(sourceRun, targetRun)
                        targetRun.setText('', 0)

                        String bookmarkName = "${name}_${number}"
                        CTBookmark bookmark = newParagraph.getCTP().addNewBookmarkStart()
                        bookmark.setName(bookmarkName)
                        bookmark.setId(BigInteger.valueOf(0))

                        CTSimpleField ctSimpleField = newParagraph.getCTP().addNewFldSimple()
                        ctSimpleField.setInstr("SEQ ${name} \\* MERGEFORMAT")

                        newParagraph.getCTP().addNewBookmarkEnd().setId(BigInteger.valueOf(0))

                        String paragraphText = sourceParagraph.getText()
                        Matcher matcher = templatePattern.matcher(paragraphText)
                        matcher.find()
                        String templateContent = matcher.group(1)
                        numberTemplateBookmarkMap.put(templateContent, bookmarkName)

                        targetRun = newParagraph.createRun()
                        copyRun(sourceRun, targetRun)
                        String newText = sourceRunText.substring(1)
                        targetRun.setText(newText, 0)
                        templateParsed = true
                    }

                    if (!templateFound) {
                        XWPFRun targetRun = newParagraph.createRun()
                        copyRun(sourceRun, targetRun)
                    }
                }
            }
            return numberTemplateBookmarkMap
        }

        private static void setRefsToNumber(IBody body, List<XWPFParagraph> paragraphs, String numberTemplate, String bookmarkName) {
            for (sourceParagraph in paragraphs) {
                XmlCursor cursor = sourceParagraph.getCTP().newCursor()
                XWPFParagraph newParagraph = body.insertNewParagraph(cursor)
                CTPPr pPr = newParagraph.getCTP().isSetPPr() ? newParagraph.getCTP().getPPr() : newParagraph.getCTP().addNewPPr()
                pPr.set(sourceParagraph.getCTP().getPPr())

                boolean templateFound = false
                boolean templateParsed = false
                XWPFRun beforeRun = null
                List<XWPFRun> templateRuns = []
                for (sourceRun in sourceParagraph.getRuns()) {
                    String sourceRunText = sourceRun.getText(0)

                    if (templateParsed) {
                        XWPFRun targetRun = newParagraph.createRun()
                        copyRun(sourceRun, targetRun)
                        continue
                    }

                    if (!templateFound && sourceRunText && sourceRunText.contains('<')) {
                        XWPFRun targetRun = newParagraph.createRun()
                        copyRun(sourceRun, targetRun)

                        beforeRun = targetRun
                        templateFound = true
                        continue
                    }

                    if (!templateParsed && sourceRunText && sourceRunText.contains('>')) {
                        List<String> templateRunTexts = templateRuns.collect { XWPFRun run -> run.getText(0) }
                        String templateText = String.join('', templateRunTexts)

                        if (templateText != numberTemplate) {
                            beforeRun = null

                            for (templateRun in templateRuns) {
                                XWPFRun targetRun = newParagraph.createRun()
                                copyRun(templateRun, targetRun)
                            }

                            XWPFRun targetRun = newParagraph.createRun()
                            copyRun(sourceRun, targetRun)

                            templateFound = false
                            continue
                        }

                        String beforeRunText = beforeRun.getText(0)
                        String newText = beforeRunText.substring(0, beforeRunText.size() - 1)
                        beforeRun.setText(newText, 0)

                        XWPFRun targetRun = newParagraph.createRun()
                        copyRun(sourceRun, targetRun)
                        targetRun.setText('', 0)

                        CTSimpleField ctSimpleField = newParagraph.getCTP().addNewFldSimple()
                        ctSimpleField.setInstr("REF ${bookmarkName} \\h")

                        targetRun = newParagraph.createRun()
                        copyRun(sourceRun, targetRun)
                        newText = sourceRunText.substring(1)
                        targetRun.setText(newText, 0)
                        templateParsed = true
                    }

                    if (!templateFound) {
                        XWPFRun targetRun = newParagraph.createRun()
                        copyRun(sourceRun, targetRun)
                    }

                    if (templateFound) {
                        templateRuns.add(sourceRun)
                    }
                }
            }
        }

        void setBusinessProcessSectionNumbers() {
            Pattern pattern = getBusinessProcessSectionPattern()
            List<XWPFParagraph> businessProcessSectionParagraphs = findParagraphsByPattern(document, pattern)

            XWPFTable table = findTableByHeaders(document, BUSINESS_PROCESS_HIERARCHY_TABLE_HEADERS)

            if (businessProcessSectionParagraphs.size() != table.getRows().size() - 3) {
                throw new Exception('Количество записей сценариев и процедур в таблице "Место Процесса в иерархии процессов" не соответсвует количеству разделов для данных записей')
            }

            for (int rowNumber = 1; rowNumber < 3; rowNumber++) {
                XWPFParagraph numberParagraph = table.getRows().get(rowNumber).getTableCells().get(2).getParagraphs().get(0)
                XWPFRun targetRun = getOnlyOneParagraphRun(numberParagraph)
                targetRun.setText('', 0)
            }

            int rowNumber = 2
            for (sourceSectionParagraph in businessProcessSectionParagraphs) {
                rowNumber += 1

                String bookmarkName = createBusinessProcessSectionBookmark(sourceSectionParagraph, rowNumber)
                removeParagraph(document, sourceSectionParagraph)

                XWPFParagraph numberParagraph = table.getRows().get(rowNumber).getTableCells().get(2).getParagraphs().get(0)
                XWPFRun targetRun = getOnlyOneParagraphRun(numberParagraph)
                targetRun.setText('', 0)

                CTSimpleField ctSimpleField = numberParagraph.getCTP().addNewFldSimple()
                ctSimpleField.setInstr("REF ${bookmarkName} \\r \\h")
            }
        }

        private Pattern getBusinessProcessSectionPattern() {
            List<String> partsPattern = []
            for (scenario in subprocessDescription.scenarios) {
                String scenarioCode = scenario.scenario.functionInfo.code ? scenario.scenario.functionInfo.code : "<${SCENARIO_CODE_TEMPLATE_KEY}>"
                String scenarioName = scenario.scenario.functionInfo.function.name ? scenario.scenario.functionInfo.function.name : "<${SCENARIO_NAME_TEMPLATE_KEY}>"
                String scenarioPattern = "Сценарий ${scenarioCode} ${scenarioName}"
                partsPattern.add(scenarioPattern)

                if (detailLevel == 4) {
                    for (procedure in scenario.procedures) {
                        String procedureCode = procedure.procedure.functionInfo.code ? procedure.procedure.functionInfo.code : "<${PROCEDURE_CODE_TEMPLATE_KEY}>"
                        String procedureName = procedure.procedure.functionInfo.function.name ? procedure.procedure.functionInfo.function.name : "<${PROCEDURE_NAME_TEMPLATE_KEY}>"
                        String procedurePattern = "Процедура ${procedureCode} ${procedureName}"
                        partsPattern.add(procedurePattern)
                    }
                }
            }

            partsPattern = partsPattern.collect { String part -> Pattern.quote(part) }
            // noinspection RegExpUnnecessaryNonCapturingGroup
            Pattern pattern = Pattern.compile("^(?:(?:${String.join(')|(?:', partsPattern)}))\$")
            return pattern
        }

        private String createBusinessProcessSectionBookmark(XWPFParagraph sourceSectionParagraph, int rowNumber) {
            XmlCursor cursor = sourceSectionParagraph.getCTP().newCursor()
            XWPFParagraph newSectionParagraph = document.insertNewParagraph(cursor)
            CTPPr pPr = newSectionParagraph.getCTP().isSetPPr() ? newSectionParagraph.getCTP().getPPr() : newSectionParagraph.getCTP().addNewPPr()
            pPr.set(sourceSectionParagraph.getCTP().getPPr())

            String bookmarkName = "Business_process_${rowNumber}"
            CTBookmark bookmark = newSectionParagraph.getCTP().addNewBookmarkStart()
            bookmark.setName(bookmarkName)
            bookmark.setId(BigInteger.valueOf(0))

            for (sourceRun in sourceSectionParagraph.getRuns()) {
                XWPFRun targetRun = newSectionParagraph.createRun()
                copyRun(sourceRun, targetRun)
            }

            newSectionParagraph.getCTP().addNewBookmarkEnd().setId(BigInteger.valueOf(0))

            return bookmarkName
        }

        void setDocumentCollectionHyperlinks() {
            XWPFTable table = findTableByHeaders(document, DOCUMENT_COLLECTION_TABLE_HEADERS)

            if (table.getRows().size() == 1) {
                return
            }

            for (int rowNumber = 1; rowNumber < table.getRows().size(); rowNumber++) {
                XWPFParagraph paragraph = table.getRow(rowNumber).getTableCells().get(0).getParagraphs().get(0)
                String paragraphText = paragraph.getText()

                if (paragraphText.contains("<${PROCESS_DOCUMENT_COLLECTION_TEMPLATE_KEY}>")) {
                    continue
                }

                String bookmarkName = createDocumentCollectionBookmark(paragraph, rowNumber)

                setHyperlinksInDocumentCollectionTable(table, paragraphText, bookmarkName)

                List<XWPFTable> functionsTables = []
                for (currentTable in document.getTables()) {
                    if (currentTable.getRows().size() == 0) {
                        continue
                    }

                    int headerMatchesCount = 0
                    for (row in currentTable.getRows()) {
                        String horizontalHeader = row.getTableCells().get(0).getParagraphs().get(0).getText()

                        if (horizontalHeader in FUNCTIONS_TABLE_DOCUMENT_HORIZONTAL_HEADERS) {
                            headerMatchesCount++
                        }
                    }

                    if (headerMatchesCount == 2) {
                        functionsTables.add(currentTable)
                    }
                }
                setHyperlinksInFunctionsTables(functionsTables, paragraphText, bookmarkName)
            }
        }

        private static String createDocumentCollectionBookmark(XWPFParagraph paragraph, int rowNumber) {
            String paragraphText = paragraph.getText()
            int delimiterSymbolIndex = paragraphText.indexOf('[') - 1
            String documentCollectionName = paragraphText.substring(0, delimiterSymbolIndex)
            String documentCollectionInfo = paragraphText.substring(delimiterSymbolIndex)

            XWPFRun sourceRun = getOnlyOneParagraphRun(paragraph)
            sourceRun.setText('', 0)

            String bookmarkName = "Document_collection_${rowNumber}"
            CTBookmark bookmark = paragraph.getCTP().addNewBookmarkStart()
            bookmark.setName(bookmarkName)
            bookmark.setId(BigInteger.valueOf(0))

            XWPFRun targetRun = paragraph.createRun()
            copyRun(sourceRun, targetRun)
            targetRun.setText(documentCollectionName, 0)

            paragraph.getCTP().addNewBookmarkEnd().setId(BigInteger.valueOf(0))

            targetRun = paragraph.createRun()
            copyRun(sourceRun, targetRun)
            targetRun.setText(documentCollectionInfo, 0)

            return bookmarkName
        }

        private static void setHyperlinksInDocumentCollectionTable(XWPFTable table, String documentCollection, String bookmarkName) {
            if (table.getRows().size() == 1) {
                return
            }

            for (int rowNumber = 1; rowNumber < table.getRows().size(); rowNumber++) {
                for (paragraph in table.getRow(rowNumber).getTableCells().get(1).getParagraphs()) {
                    if (!(paragraph.getText().contains(documentCollection))) {
                        continue
                    }

                    setHyperlinkToDocumentCollection(paragraph, bookmarkName)
                }
            }
        }

        private static void setHyperlinksInFunctionsTables(List<XWPFTable> tables, String documentCollection, String bookmarkName) {
            for (table in tables) {
                for (row in table.getRows()) {
                    String horizontalHeader = row.getTableCells().get(0).getParagraphs().get(0).getText()

                    if (horizontalHeader in FUNCTIONS_TABLE_DOCUMENT_HORIZONTAL_HEADERS) {
                        XWPFTableCell cell = row.getTableCells().get(1)
                        for (paragraph in cell.getParagraphs()) {
                            if (!(paragraph.getText().contains(documentCollection))) {
                                continue
                            }

                            setHyperlinkToDocumentCollection(paragraph, bookmarkName)
                        }
                    }
                }
            }
        }

        private static void setHyperlinkToDocumentCollection(XWPFParagraph paragraph, String bookmarkName) {
            String paragraphText = paragraph.getText()
            int delimiterSymbolIndex = paragraphText.indexOf('[') - 1
            String documentCollectionName = paragraphText.substring(0, delimiterSymbolIndex)
            String documentCollectionInfo = paragraphText.substring(delimiterSymbolIndex)

            XWPFRun sourceRun = getOnlyOneParagraphRun(paragraph)
            sourceRun.setText('', 0)

            CTHyperlink hyperlink = paragraph.getCTP().addNewHyperlink()
            hyperlink.setAnchor(bookmarkName)
            hyperlink.addNewR()

            XWPFHyperlinkRun hyperlinkRun = new XWPFHyperlinkRun(hyperlink, hyperlink.getRArray(0), paragraph)
            hyperlinkRun.setText(documentCollectionName)
            hyperlinkRun.setColor('0000FF')
            hyperlinkRun.setUnderline(UnderlinePatterns.SINGLE)

            XWPFRun targetRun = paragraph.createRun()
            copyRun(sourceRun, targetRun)
            targetRun.setText(documentCollectionInfo, 0)
        }

        private void replaceHeadersText(String pattern, String replacement) {
            for (header in document.getHeaderList()) {
                for (headerParagraph in header.getParagraphs()) {
                    replaceParagraphText(headerParagraph, pattern, replacement)
                }
            }
        }

        private static void replaceInCopyParagraph(IBody body, String pattern, String replacement) {
            if (pattern == replacement) {
                return
            }

            if (pattern == replacement.substring(0, replacement.size() - 1)) {
                return
            }

            List<XWPFParagraph> paragraphs = findParagraphsByText(body, pattern)
            paragraphs.each { XWPFParagraph paragraph ->
                XWPFParagraph newParagraph = addParagraph(body, paragraph)
                replaceParagraphText(newParagraph, pattern, replacement)
            }
        }

        private void replaceParagraphsText(String pattern, String replacement) {
            for (paragraph in document.getParagraphs()) {
                replaceParagraphText(paragraph, pattern, replacement)
            }
        }

        private static void replaceParagraphsText(List<IBodyElement> elements, String pattern, String replacement) {
            for (element in elements) {
                if (element.getElementType() == BodyElementType.PARAGRAPH) {
                    replaceParagraphText((XWPFParagraph) element, pattern, replacement)
                }
            }
        }

        private static void replaceParagraphText(XWPFParagraph paragraph, String pattern, String replacement) {
            if (paragraph.getText().contains(pattern)) {
                String newText = paragraph.getText().replace(pattern, replacement)
                addParagraphText(paragraph, newText)
            }
        }

        private static List<IBodyElement> findBodyElements(XWPFDocument document, String startParagraphTextPart, String stopParagraphTextPart, int escapeCount = 0) {
            List<IBodyElement> elements = []
            for (bodyElement in document.getBodyElements()) {
                if (!elements && bodyElement instanceof XWPFParagraph) {
                    XWPFParagraph paragraph = (XWPFParagraph) bodyElement
                    if (paragraph.getText().contains(startParagraphTextPart)) {
                        if (escapeCount == 0) {
                            elements.add(bodyElement)
                            continue
                        } else {
                            escapeCount -= 1
                        }
                    }
                }

                if (elements && bodyElement instanceof XWPFParagraph) {
                    XWPFParagraph paragraph = (XWPFParagraph) bodyElement
                    if (paragraph.getText().contains(stopParagraphTextPart)) {
                        break
                    }
                }

                if (elements) {
                    elements.add(bodyElement)
                }
            }
            return elements
        }

        private static List<XWPFParagraph> findParagraphsByPattern(IBody body, Pattern pattern) {
            List<XWPFParagraph> foundedParagraphs = []
            for (paragraph in body.getParagraphs()) {
                String paragraphText = paragraph.getText()
                Matcher matcher = pattern.matcher(paragraphText)

                if (matcher.find()) {
                    foundedParagraphs.add(paragraph)
                }
            }
            return foundedParagraphs
        }

        private static List<XWPFParagraph> findParagraphsByText(IBody body, String text) {
            return body.getParagraphs()
                    .findAll { XWPFParagraph paragraph -> paragraph.getText().contains(text) }
        }

        private static XWPFTable findTableByHeaders(XWPFDocument document, List<String> tableHeaders) {
            XWPFTable foundedTable = null
            for (table in document.getTables()) {
                boolean tableFound = tableHasHeaders(table, tableHeaders)

                if (tableFound) {
                    foundedTable = table
                    break
                }
            }
            return foundedTable
        }

        private static XWPFTable findTableByHeaders(List<IBodyElement> elements, List<String> tableHeaders) {
            XWPFTable foundedTable = null
            for (element in elements) {
                if (element.getElementType() == BodyElementType.TABLE) {
                    XWPFTable table = (XWPFTable) element
                    boolean tableFound = tableHasHeaders(table, tableHeaders)

                    if (tableFound) {
                        foundedTable = table
                        break
                    }
                }
            }
            return foundedTable
        }

        private static boolean listHasNotPatternValue(List<String> list, String pattern) {
            for (value in list) {
                if (value != pattern) {
                    return true
                }
            }

            return false
        }

        private static boolean tableHasHeaders(XWPFTable table, List<String> headers) {
            if (table.getRows().size() == 0) {
                return false
            }

            if (table.getRows().get(0).getTableCells().size() != headers.size()) {
                return false
            }

            for (int columnNumber = 0; columnNumber < headers.size(); columnNumber++) {
                String header = headers[columnNumber]
                XWPFTableCell cell = table.getRows().get(0).getTableCells().get(columnNumber)

                if (cell.getParagraphs().size() == 0) {
                    return false
                }

                XWPFParagraph paragraph = cell.getParagraphs().get(0)
                if (paragraph.getText() != header) {
                    return false
                }
            }

            return true
        }

        private XWPFParagraph addPicture(XWPFParagraph imageParagraph, Model model, XWPFParagraph labelParagraph) {
            try {
                ModelImage modelImage = getModelImage(model, imageType)

                if (modelImage.image.length == 0) {
                    throw new Exception("Изображение не найдено")
                }

                XWPFRun run = imageParagraph.createRun()
                imageParagraph.setAlignment(ParagraphAlignment.CENTER)

                // Если изображение широкое, то разворачиваем страницу
                if (modelImage.width > modelImage.height) {
                    // ориентацию настраиваем для последнего параграфа на странице (надписи под рисунком), так как
                    // после данного параграфа автоматически ставится разрыв раздела
                    setPageOrientation(labelParagraph, STPageOrientation.LANDSCAPE)
                } else {
                    // ориентацию настраиваем для последнего параграфа на странице (надписи под рисунком), так как
                    // после данного параграфа автоматически ставится разрыв раздела
                    setPageOrientation(labelParagraph, STPageOrientation.PORTRAIT)
                }

                int shortSide
                int longSide

                if (modelImage.type == ImageType.PNG) {
                    shortSide = 450
                    longSide = 700
                } else if (modelImage.type == ImageType.SVG) {
                    final Double A4WCM = 21.0D
                    final Double A4HCM = 29.7D

                    shortSide = getCmInPoints(A4WCM)
                    longSide = getCmInPoints(A4HCM)
                } else {
                    throw new Exception('Неподдерживаемый тип изображения')
                }

                final int labelStringHeight = 16
                int labelLength = labelParagraph.getText().size()

                int pageW
                int pageH

                if (modelImage.width > modelImage.height) {
                    int labelStringsCount = (int) Math.ceil(labelLength / 95.0) // 95 букв в одной строке
                    pageW = longSide
                    pageH = shortSide - labelStringsCount * labelStringHeight
                } else {
                    int labelStringsCount = (int) Math.ceil(labelLength / 57.0) // 57 букв в одной строке
                    pageW = shortSide
                    pageH = longSide - labelStringsCount * labelStringHeight
                }

                // Если изображение не помещается на страницу, то надо его масштабировать
                double scale = 1.0
                if (modelImage.width > pageW || modelImage.height > pageH) {
                    double widthScale = pageW / modelImage.width
                    double heightScale = pageH / modelImage.height
                    scale = Math.min(widthScale, heightScale)
                }

                int width = (int) (modelImage.width * scale)
                int height = (int) (modelImage.height * scale)

                modelImage.addToRun(document, run, width, height)
            }
            catch (Exception e) {
                setPageOrientation(labelParagraph, STPageOrientation.PORTRAIT)
                addParagraphText(imageParagraph, "Ошибка вставки изображения: ${e.getMessage()}")
            }

            return imageParagraph
        }

        private static void setPageOrientation(XWPFParagraph paragraph, STPageOrientation.Enum orientation) {
            CTSectPr sect = paragraph.getCTPPr().getSectPr() ? paragraph.getCTPPr().getSectPr() : paragraph.getCTPPr().addNewSectPr()
            CTPageSz pageSize = sect.getPgSz() ? sect.getPgSz() : sect.addNewPgSz()
            pageSize.setOrient(orientation)
            if (orientation == STPageOrientation.LANDSCAPE) {
                pageSize.setW(842 * 20)
                pageSize.setH(595 * 20)
            } else {
                pageSize.setW(595 * 20)
                pageSize.setH(842 * 20)
            }
        }

        private static int getCmInPoints(Double cm) {
            return (int) (cm * 28.3464567)
        }

        private static XWPFParagraph addParagraph(IBody body, XWPFParagraph paragraph) {
            XmlCursor cursor = paragraph.getCTP().newCursor()
            XWPFParagraph newParagraph = body.insertNewParagraph(cursor)
            copyParagraph(paragraph, newParagraph)
            return newParagraph
        }

        private static void addParagraphText(XWPFParagraph paragraph, String text, int fontSize = -1) {
            XWPFRun run = getOnlyOneParagraphRun(paragraph)

            if (text.contains('""')) {
                text = text.replaceAll('""', '"')
            }

            while (text.startsWith("\\")) {
                text = parseStyleSymbols(text, run, paragraph)
            }

            if (fontSize > 0) {
                run.setFontSize(fontSize)
            }

            List<String> textParts = []
            String currentTextPart = ''
            for (int symbolPosition = 0; symbolPosition < text.size(); symbolPosition++) {
                if (text[symbolPosition] == '<') {
                    currentTextPart += text[symbolPosition]
                    textParts.add(currentTextPart)
                    currentTextPart = ''
                    continue
                }

                if (text[symbolPosition] == '>') {
                    textParts.add(currentTextPart)
                    currentTextPart = text[symbolPosition]
                    continue
                }

                currentTextPart += text[symbolPosition]
            }

            if (currentTextPart) {
                textParts.add(currentTextPart)
            }

            run.setText(textParts[0], 0)
            paragraph.addRun(run)

            int partNumber = 1
            for (textPart in textParts) {
                if (partNumber == 1) {
                    partNumber++
                    continue
                }

                XWPFRun currentRun = paragraph.createRun()
                copyRun(run, currentRun)

                if (partNumber % 2 == 0) {
                    CTRPr rPr = currentRun.getCTR().isSetRPr() ? currentRun.getCTR().getRPr() : currentRun.getCTR().addNewRPr()
                    CTShd cTShd = rPr.addNewShd()
                    cTShd.setFill("C0C0C0")
                }

                currentRun.setText(textPart, 0)
                paragraph.addRun(currentRun)

                partNumber++
            }
        }

        private static List<IBodyElement> copyIBodyElements(List<IBodyElement> elements, IBody targetBody = null, XmlCursor cursor = null) {
            if (targetBody == null || cursor == null) {
                IBodyElement lastElement = elements.get(elements.size() - 1)

                if (targetBody == null)
                    targetBody = lastElement.getBody()

                if (cursor == null) {
                    cursor = lastElement.getElementType() == BodyElementType.PARAGRAPH ? ((XWPFParagraph) lastElement).getCTP().newCursor() : ((XWPFTable) lastElement).getCTTbl().newCursor()
                    cursor.toNextSibling()
                }
            }

            List<IBodyElement> newElements = []
            elements.each { IBodyElement element ->
                newElements.add(copyIBodyElement(element, targetBody, cursor))
            }
            return newElements
        }

        private static IBodyElement copyIBodyElement(IBodyElement element, IBody targetBody = null, XmlCursor cursor = null) {
            IBodyElement newElement = null

            if (element.getElementType() == BodyElementType.PARAGRAPH) {
                if (targetBody == null || cursor == null) {
                    if (targetBody == null)
                        targetBody = element.getBody()

                    if (cursor == null) {
                        cursor = ((XWPFParagraph) element).getCTP().newCursor()
                        cursor.toNextSibling()
                    }
                }

                newElement = targetBody.insertNewParagraph(cursor)
                copyParagraph((XWPFParagraph) element, (XWPFParagraph) newElement)
            }

            if (element.getElementType() == BodyElementType.TABLE) {
                if (targetBody == null || cursor == null) {
                    targetBody = element.getBody()
                    cursor = ((XWPFTable) element).getCTTbl().newCursor()
                }

                newElement = targetBody.insertNewTbl(cursor)
                copyTable((XWPFTable) element, (XWPFTable) newElement)
            }

            cursor.toNextToken()
            return newElement
        }

        private static void copyTable(XWPFTable source, XWPFTable target) {
            target.getCTTbl().setTblPr(source.getCTTbl().getTblPr())
            target.getCTTbl().setTblGrid(source.getCTTbl().getTblGrid())

            // по умолчанию в таблице создается одна строка, удаляем её
            target.removeRow(0)

            for (int rowNum = 0; rowNum < source.getRows().size(); rowNum++) {
                XWPFTableRow sourceRow = source.getRows().get(rowNum)
                copyTableRow(sourceRow, target)
            }
        }

        private static XWPFTableRow copyTableRow(XWPFTableRow sourceRow, XWPFTable table) {
            XWPFTableRow newRow = table.createRow()

            while (newRow.getTableCells().size() > sourceRow.getTableCells().size()) {
                newRow.removeCell(newRow.getTableCells().size() - 1)
            }

            while (newRow.getTableCells().size() < sourceRow.getTableCells().size()) {
                newRow.createCell()
            }

            newRow.getCtRow().setTrPr(sourceRow.getCtRow().getTrPr())

            for (int cellNumber = 0; cellNumber < sourceRow.getTableCells().size(); cellNumber++) {
                XWPFTableCell targetCell = newRow.getTableCells().get(cellNumber)
                XWPFTableCell sourceCell = sourceRow.getTableCells().get(cellNumber)
                copyTableCell(sourceCell, targetCell)
            }
            return newRow
        }

        private static XWPFTableCell copyTableCell(XWPFTableCell sourceCell, XWPFTableCell targetCell) {
            targetCell.getCTTc().setTcPr(sourceCell.getCTTc().getTcPr())

            for (int paragraphNumber = 0; paragraphNumber < sourceCell.getParagraphs().size(); paragraphNumber++) {
                XWPFParagraph sourceParagraph = sourceCell.getParagraphs().get(paragraphNumber)
                XWPFParagraph targetParagraph = targetCell.getParagraphs().get(paragraphNumber)
                copyParagraph(sourceParagraph, targetParagraph)
            }
            return targetCell
        }

        private static void copyParagraph(XWPFParagraph source, XWPFParagraph target) {
            CTPPr pPr = target.getCTP().isSetPPr() ? target.getCTP().getPPr() : target.getCTP().addNewPPr()
            pPr.set(source.getCTP().getPPr())
            for (sourceRun in source.getRuns()) {
                XWPFRun targetRun = target.createRun()
                copyRun(sourceRun, targetRun)
            }
        }

        private static void copyRun(XWPFRun source, XWPFRun target) {
            CTRPr rPr = target.getCTR().isSetRPr() ? target.getCTR().getRPr() : target.getCTR().addNewRPr()
            rPr.set(source.getCTR().getRPr())
            target.setText(source.getText(0))
        }

        private static void removeBodyElements(XWPFDocument document, List<IBodyElement> elementsToDelete) {
            elementsToDelete.each { IBodyElement element ->
                int elementPosition = -1

                if (element instanceof XWPFParagraph) {
                    XWPFParagraph paragraph = (XWPFParagraph) element
                    elementPosition = document.getPosOfParagraph(paragraph)
                }

                if (element instanceof XWPFTable) {
                    XWPFTable table = (XWPFTable) element
                    elementPosition = document.getPosOfTable(table)
                }

                document.removeBodyElement(elementPosition)
            }
        }

        private static void removeParagraph(XWPFDocument document, XWPFParagraph paragraph) {
            int position = document.getPosOfParagraph(paragraph)
            document.removeBodyElement(position)
        }

        private static XWPFRun getOnlyOneParagraphRun(XWPFParagraph paragraph) {
            if (paragraph.getRuns()) {
                while (paragraph.getRuns().size() > 1) {
                    paragraph.removeRun(1)
                }
                return paragraph.getRuns().get(0)
            }

            return paragraph.createRun()
        }

        private static String parseStyleSymbols(String text, XWPFRun run, XWPFParagraph paragraph) {
            if (text.startsWith("\\B")) {
                run.setBold(true)
                String newText = text.substring(2)
                return newText
            }

            if (text.startsWith("\\!B")) {
                run.setBold(false)
                return text.substring(3)
            }

            if (text.startsWith("\\I")) {
                run.setItalic(true)
                return text.substring(2)
            }

            if (text.startsWith("\\U")) {
                run.setUnderline(UnderlinePatterns.SINGLE)
                return text.substring(2)
            }

            if (text.startsWith("\\L")) {
                text = text.substring(2)

                try {
                    String level = text.substring(0, 1)
                    paragraph.setIndentationLeft(level.toInteger() * CM_1_OFFSET)
                    return text.substring(1)
                }
                catch (Exception ignored) {
                    paragraph.setIndentationLeft(CM_1_OFFSET)
                    return text
                }
            }

            return text.substring(1)
        }

        void enforceUpdateFields() {
            document.enforceUpdateFields()
        }

        void saveContent() {
            ByteArrayOutputStream outputStream = new ByteArrayOutputStream()
            document.write(outputStream)
            byte[] bytes = outputStream.toByteArray()
            long userId = context.principalId()

            content = FileNodeDTO.builder()
                    .nodeId(NodeId.builder().id(UUID.randomUUID().toString()).repositoryId(FILE_REPOSITORY_ID).build())
                    .parentNodeId(NodeId.builder().id(String.valueOf(userId)).repositoryId(FILE_REPOSITORY_ID).build())
                    .extension(DOCX_FORMAT)
                    .file(new SimpleMultipartFile(fileName, bytes))
                    .name(fileName + '.' + DOCX_FORMAT)
                    .build()
        }
    }

    private static String getName(TreeNode treeNode, boolean onlyShortName = false) {
        Node node = treeNode._getNode()
        String name = node.getName()
        name = name ? trimStringValue(name) : ''
        if (name) {
            findAbbreviations(name)
        }

        if (onlyShortName) {
            return name
        }

        String fullName = getAttributeValue(treeNode, FULL_NAME_ATTR_ID)
        if (fullName) {
            findAbbreviations(fullName)
        }

        return fullName ? fullName : name
    }

    private static String getAttributeValue(TreeNode treeNode, String attributeId) {
        Node node = treeNode._getNode()
        AttributeValue attribute = node.getAttributes().stream()
                .filter { AttributeValue aV -> aV.typeId == attributeId }
                .findFirst()
                .orElse(null)

        if (attribute != null && attribute.value != null && !attribute.value.trim().isEmpty()) {
            String value = trimStringValue(attribute.value)

            if (value) {
                findAbbreviations(value)
            }

            return value
        }

        return ''
    }

    private static String trimStringValue(String value) {
        String resultString = value.replaceAll('\\u00A0', ' ')
        resultString = resultString.replaceAll('[\\s\\n]+', ' ').trim()
        return resultString
    }

    private static void findAbbreviations(String value) {
        Matcher matcher = abbreviationsPattern.matcher(value)
        while (matcher.find()) {
            String abbreviationName = value.substring(matcher.start(), matcher.end())

            if (abbreviationName in foundedAbbreviations.keySet()) {
                continue
            }

            String abbreviationDescription = fullAbbreviations.get(abbreviationName)
            foundedAbbreviations.put(abbreviationName, abbreviationDescription)
        }
    }

    private static Model getEPCModel(ObjectElement objectElement) {
        return objectElement.getObjectDefinition()
                .getDecompositions(EPC_MODEL_TYPE_ID)
                .stream()
                .findFirst()
                .orElse(null)
    }

    private ModelImage getModelImage(Model model, ImageType type) {
        if (type == ImageType.PNG) {
            byte[] image = model.getImagePng()
            InputStream is = new ByteArrayInputStream(image)
            BufferedImage img = ImageIO.read(is)
            return new ModelImage(image, type, img.width, img.height)
        }

        if (type == ImageType.SVG) {
            FullModelDefinition modelDefinition = modelApi.getFullModelDefinition(model.getRepositoryId(), model.getId())
            String svg = imageApi.getImageSvg(modelDefinition)
            svg = svg.replaceAll('\n', '')
            svg = svg.replaceAll(' +', ' ')

            byte[] image = svg.getBytes()
            int width = extractWidthSvgImage(svg)
            int height = extractHeightSvgImage(svg)
            return new ModelImage(image, type, width, height)
        }

        throw new Exception('Неподдерживаемый тип изображения')
    }

    private static int extractWidthSvgImage(String svg) {
        int firstIndex = svg.indexOf('width=')
        int secondIndex = svg.indexOf('"', firstIndex) + 1
        int thirdIndex = svg.indexOf('"', secondIndex)
        String width = svg.substring(secondIndex, thirdIndex)
        return Double.parseDouble(width)
    }

    private static int extractHeightSvgImage(String svg) {
        int firstIndex = svg.indexOf('height=')
        int secondIndex = svg.indexOf('"', firstIndex) + 1
        int thirdIndex = svg.indexOf('"', secondIndex)
        String height = svg.substring(secondIndex, thirdIndex)
        return Double.parseDouble(height)
    }

    @Override
    void execute() {
        init()

        List<ObjectElement> subprocessObjects = getSubprocessObjects()
        List<SubprocessDescription> subprocessDescriptions = getSubProcessDescriptions(subprocessObjects)
        List<FileInfo> files = createBPRegulations(subprocessDescriptions)

        String resultFileName = null
        String format = null
        FileNodeDTO result = null

        if (files.size() == 1) {
            resultFileName = files[0].name
            format = DOCX_FORMAT
            result = files[0].content
        }

        if (files.size() > 1) {
            resultFileName = ZIP_RESULT_FILE_NAME_FIRST_PART + ' ' + detailLevel.toString() + '_' + new SimpleDateFormat('yyyyMMdd HHmmss').format(new Date()).replace(' ', '_')
            format = ZIP_FORMAT

            byte[] zipFileContent = createZipFileContent(files)
            long userId = context.principalId()

            FileNodeDTO fileNode = FileNodeDTO.builder()
                    .nodeId(NodeId.builder().id(UUID.randomUUID().toString()).repositoryId(FILE_REPOSITORY_ID).build())
                    .parentNodeId(NodeId.builder().id(String.valueOf(userId)).repositoryId(FILE_REPOSITORY_ID).build())
                    .extension(format)
                    .file(new SimpleMultipartFile(resultFileName, zipFileContent))
                    .name(resultFileName + '.' + format)
                    .build()
            result = fileNode
        }

        if (result == null) {
            return
        }

        if (!DEBUG) {
            context.getApi(FileApi).uploadFile(result)
        }

        context.setResultFile(result.file.bytes, format, resultFileName)
    }

    private void init() {
        imageApi = context.getApi(ImageApi.class)
        modelApi = context.getApi(ModelApi.class)
        treeRepository = context.createTreeRepository(true)
        parseParameters()
        initAbbreviations()
    }

    private void parseParameters() {
        if (DEBUG) {
            detailLevel = 4
            docVersion = '01'
            docDate = '01.01.2025'
            imageType = ImageType.PNG
            return
        }

        String deep = ParamUtils.parse(context.findParameter(DETAIL_LEVEL_PARAM_NAME)) as String
        detailLevel = Integer.parseInt(deep.replaceAll('[^0-9]+', ''))

        String docVersionParam = ParamUtils.parse(context.findParameter(DOC_VERSION_PARAM_NAME)) as String
        if (docVersionParam) {
            docVersion = docVersionParam
        } else {
            docVersion = ''
        }

        String docDateParam = ParamUtils.parse(context.findParameter(DOC_DATE_PARAM_NAME)) as String
        if (docDateParam) {
            Timestamp approvalDate = ParamUtils.parse(context.findParameter(DOC_DATE_PARAM_NAME)) as Timestamp
            docDate = approvalDate.format('dd.MM.yyyy')
        } else {
            docDate = ''
        }

        String imgType = ParamUtils.parse(context.findParameter(IMAGE_TYPE_PARAM_NAME)) as String
        if (imgType == 'PNG') {
            imageType = ImageType.PNG
        } else if (imgType == 'SVG') {
            imageType = ImageType.SVG
        } else {
            throw new Exception('Неподдерживаемый формат изображения')
        }
    }

    private void initAbbreviations() {
        Model abbreviationsModel = treeRepository.read(context.modelId().getRepositoryId(), ABBREVIATIONS_MODEL_ID)
        if (!abbreviationsModel) {
            throw new SilaScriptException("Неверный ID модели аббревиатур [${ABBREVIATIONS_MODEL_ID}]")
        }

        ObjectElement abbreviationsRootObject = abbreviationsModel.getObjects()
                .find { ObjectElement oE -> oE.getObjectDefinitionId() == ABBREVIATIONS_ROOT_OBJECT_ID }

        if (!abbreviationsRootObject) {
            throw new SilaScriptException("Неверный ID корневого объекта аббревиатур [${ABBREVIATIONS_ROOT_OBJECT_ID}]")
        }

        List<ObjectElement> abbreviationObjects = abbreviationsRootObject.getExitEdges()
                .findAll { Edge e -> e.getEdgeTypeId() in ABBREVIATION_EDGE_TYPE_IDS }
                .collect { Edge e -> e.getTarget() as ObjectElement }
                .unique(Comparator.comparing { ObjectElement oE -> oE.getId() })

        abbreviationObjects.addAll(
                abbreviationsRootObject.getEnterEdges()
                        .findAll { Edge e -> e.getEdgeTypeId() in ABBREVIATION_EDGE_TYPE_IDS }
                        .collect { Edge e -> e.getSource() as ObjectElement }
                        .unique(Comparator.comparing { ObjectElement oE -> oE.getId() })
        )

        for (abbreviationObject in abbreviationObjects) {
            ObjectDefinitionNode abbreviationObjectDefinitionNode = abbreviationObject.getObjectDefinition()._getNode() as ObjectDefinitionNode

            String abbreviationName = abbreviationObjectDefinitionNode.getName()
            String abbreviationDescription = ''
            AttributeValue descriptionDefinitionAttribute = abbreviationObjectDefinitionNode.getAttributes().stream()
                    .filter { AttributeValue aV -> aV.typeId == DESCRIPTION_DEFINITION_ATTR_ID }
                    .findFirst()
                    .orElse(null)
            if (descriptionDefinitionAttribute != null && descriptionDefinitionAttribute.value != null && !descriptionDefinitionAttribute.value.trim().isEmpty()) {
                abbreviationDescription = descriptionDefinitionAttribute.value
            }
            fullAbbreviations.put(abbreviationName, abbreviationDescription)
        }

        Set<String> abbreviationNames = fullAbbreviations.keySet()
        // noinspection RegExpUnnecessaryNonCapturingGroup
        abbreviationsPattern = Pattern.compile("\\b(?:(?:${String.join(')|(?:', abbreviationNames)}))\\b")
    }

    private List<ObjectElement> getSubprocessObjects() {
        List<ObjectElement> subprocessObjects = []
        if (!context.elementsIdsList().isEmpty()) {
            Model model = treeRepository.read(context.modelId().getRepositoryId(), context.modelId().getId())
            List<ObjectElement> allObjects = model.getObjects()
            for (elementId in context.elementsIdsList()) {
                for (object in allObjects) {
                    if (object.getId() == elementId) {
                        subprocessObjects.add(object)
                        break
                    }
                }
            }
            subprocessObjects.sort { ObjectElement oE1, ObjectElement oE2 -> ModelUtils.getElementsCoordinatesComparator().compare(oE1, oE2) }
        }
        if (subprocessObjects.isEmpty()) {
            throw new SilaScriptException('Скрипт должен запускаться на экземплярах объектов')
        }
        return subprocessObjects
    }

    private List<SubprocessDescription> getSubProcessDescriptions(List<ObjectElement> subprocessObjects) {
        List<SubprocessDescription> subprocessDescriptions = subprocessObjects.collect { ObjectElement subprocessObject -> new SubprocessDescription(subprocessObject, detailLevel) }
        subprocessDescriptions.each { SubprocessDescription subprocessDescription -> buildSubProcessDescription(subprocessDescription) }
        return subprocessDescriptions
    }

    private void buildSubProcessDescription(SubprocessDescription subprocess) {
        subprocess.defineParentProcess()
        subprocess.findOwners()
        subprocess.defineGoals()
        subprocess.findExternalProcessInputFlows()
        subprocess.findExternalProcessOutputFlows()
        subprocess.completeExternalProcessesWithInputFlows()
        subprocess.completeExternalProcessesWithOutputFlows()
        subprocess.defineProcessSelectionModel()
        subprocess.defineScenarios()

        if (detailLevel == 4) {
            subprocess.defineProcedures()
            subprocess.defineBusinessRoles()
            subprocess.completeBusinessRoles()
            subprocess.buildResponsibilityScenariosMatrix()
        }

        subprocess.identifyAnalyzedEPC()
        subprocess.defineDocumentCollections()
        subprocess.completeDocumentCollections()
        subprocess.defineNormativeDocuments()
        subprocess.completeNormativeDocuments()
        subprocess.analyzeEPCModels()
    }

    private List<FileInfo> createBPRegulations(List<SubprocessDescription> subprocessDescriptions) {
        List<FileInfo> files = []
        subprocessDescriptions.forEach { SubprocessDescription subprocessDescription ->
            String fileName = DOCX_RESULT_FILE_NAME_FIRST_PART + " '${subprocessDescription.subprocess.function.name}' " + detailLevel.toString() + '_' + new SimpleDateFormat('yyyyMMdd HHmmss').format(new Date()).replace(' ', '_')
            BusinessProcessRegulationDocument document = getBPRegulationDocument(fileName, subprocessDescription)

            if (DEBUG) {
                FileOutputStream file = new FileOutputStream("${TEMPLATE_LOCAL_PATH}\\${fileName}.${DOCX_FORMAT}")
                document.document.write(file)
                file.close()
            }

            files.add(new FileInfo(document.fileName, document.content))
        }
        return files
    }

    private BusinessProcessRegulationDocument getBPRegulationDocument(String fileName, SubprocessDescription subprocessDescription) {
        XWPFDocument template = getTemplate()
        BusinessProcessRegulationDocument document = new BusinessProcessRegulationDocument(fileName, subprocessDescription, template, detailLevel)
        document.fillSimpleTexts()
        document.fillLists()
        document.fillTables()
        document.fillModels()
        document.setLabelNumbers()
        document.setBusinessProcessSectionNumbers()
        document.setDocumentCollectionHyperlinks()
        document.enforceUpdateFields()
        document.saveContent()
        return document
    }

    private XWPFDocument getTemplate() {
        if (DEBUG) {
            String filePath = "${TEMPLATE_LOCAL_PATH}\\${BUSINESS_PROCESS_REGULATION_TEMPLATE_NAME}"
            File file = new File(filePath)

            if (!file.exists()) {
                throw new IOException("Файл ${filePath} не найден")
            }

            try {
                FileInputStream fileInputStream = new FileInputStream(file)
                return new XWPFDocumentSvg(fileInputStream)
            } catch (Exception e) {
                log.error('Ошибка чтения файла', e)
            }
        }

        TreeNode fileFolderTreeNode = context.createTreeRepository(false).read(FILE_REPOSITORY_ID, FILE_REPOSITORY_ID)
        List<TreeNode> children = fileFolderTreeNode.getChildren()

        TreeNode fileTreeNode = null
        for (child in children) {
            if (child.getType().name() == FILE_NODE_TYPE_ID && child.getName() == TEMPLATE_FOLDER_NAME) {
                List<TreeNode> files = child.getChildren()
                for (file in files) {
                    if (file.getName().toLowerCase() == BUSINESS_PROCESS_REGULATION_TEMPLATE_NAME.toLowerCase()) {
                        fileTreeNode = file
                        break
                    }
                }
                if (fileTreeNode != null) {
                    break
                }
            }
        }

        byte[] file = context.getApi(FileApi.class).downloadFile(FILE_REPOSITORY_ID, fileTreeNode.id)
        return new XWPFDocumentSvg(new ByteArrayInputStream(file))
    }

    private byte[] createZipFileContent(List<FileInfo> files) {
        ByteArrayOutputStream byteArrayOutputStream = new ByteArrayOutputStream()
        ZipOutputStream zipOutputStream = new ZipOutputStream(byteArrayOutputStream)

        files.each { FileInfo file ->
            ZipEntry zipEntry = new ZipEntry(file.name + '.docx')
            zipOutputStream.putNextEntry(zipEntry)
            zipOutputStream.write(file.content.file.bytes, 0, file.content.file.bytes.length)
            zipOutputStream.closeEntry()
        }

        zipOutputStream.close()
        byteArrayOutputStream.close()

        return byteArrayOutputStream.toByteArray()
    }
}

class XWPFDocumentSvg extends XWPFDocument {
    XWPFDocumentSvg(InputStream is) throws IOException {
        super(is)
    }

    String addSVGPicture(InputStream is) throws Exception {
        int pictureCount = getAllPictures().size()
        RelationPart relationPart = createRelationship(XWPFRelationSvg.IMAGE_SVG, XWPFFactory.getInstance(), pictureCount + 1, false)
        XWPFPictureData img = relationPart.getDocumentPart()

        try (OutputStream out = img.getPackagePart().getOutputStream()) {
            IOUtils.copy(is, out)
        }
        pictures.add(img)
        String relationId = getRelationId(img)
        return relationId
    }
}

class XWPFRelationSvg extends POIXMLRelation {
    protected static final XWPFRelationSvg IMAGE_SVG = new XWPFRelationSvg(
            'image/svg',
            PackageRelationshipTypes.IMAGE_PART,
            '/word/media/image#.svg',
            XWPFPictureDataSvg::new,
            XWPFPictureDataSvg::new,
    )

    protected XWPFRelationSvg(
            String type,
            String rel,
            String defaultName,
            NoArgConstructor noArgConstructor,
            PackagePartConstructor packagePartConstructor
    ) {
        super(type, rel, defaultName, noArgConstructor, packagePartConstructor, null)
    }
}

class XWPFPictureDataSvg extends XWPFPictureData {
    protected XWPFPictureDataSvg() {
        super()
    }

    protected XWPFPictureDataSvg(PackagePart part) {
        super(part)
    }
}
