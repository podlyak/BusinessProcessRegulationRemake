import groovy.util.logging.Slf4j
import org.apache.poi.xwpf.usermodel.IBody
import org.apache.poi.xwpf.usermodel.UnderlinePatterns
import org.apache.poi.xwpf.usermodel.XWPFDocument
import org.apache.poi.xwpf.usermodel.XWPFParagraph
import org.apache.poi.xwpf.usermodel.XWPFRun
import org.apache.poi.xwpf.usermodel.XWPFTable
import org.apache.poi.xwpf.usermodel.XWPFTableCell
import org.apache.poi.xwpf.usermodel.XWPFTableRow
import org.apache.xmlbeans.XmlCursor
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTPPr
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTRPr
import ru.nextconsulting.bpm.dto.NodeId
import ru.nextconsulting.bpm.dto.SimpleMultipartFile
import ru.nextconsulting.bpm.repository.business.AttributeValue
import ru.nextconsulting.bpm.repository.structure.FileNodeDTO
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
import ru.nextconsulting.bpm.scriptengine.exception.SilaScriptException
import ru.nextconsulting.bpm.scriptengine.script.GroovyScript
import ru.nextconsulting.bpm.scriptengine.serverapi.FileApi
import ru.nextconsulting.bpm.scriptengine.util.ParamUtils
import ru.nextconsulting.bpm.scriptengine.util.SilaScriptParameter
import ru.nextconsulting.bpm.scriptengine.util.SilaScriptParameters
import ru.nextconsulting.bpm.utils.JsonConverter

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
                required = true
        ),
        @SilaScriptParameter(
                name = DOC_DATE_PARAM_NAME,
                type = SilaScriptParamType.DATE,
                required = true
        ),
])
@Slf4j
class BusinessProcessRegulationRemakeScript implements GroovyScript {
    static void main(String[] args) {
        ContextParameters parameters = ContextParameters.builder()
                .login('superadmin')
                .password('WM_Sila_123')
                .apiBaseUrl('http://localhost:8080/')
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

    //------------------------------------------------------------------------------------------------------------------
    // константы для работы с файлами
    //------------------------------------------------------------------------------------------------------------------
    private static final String DOCX_RESULT_FILE_NAME_FIRST_PART = 'Регламент бизнес-процесса'
    private static final String ZIP_RESULT_FILE_NAME_FIRST_PART = 'Регламенты бизнес-процессов'
    private static final String DOCX_FORMAT = 'docx'
    private static final String ZIP_FORMAT = 'zip'
    private static final String BUSINESS_PROCESS_REGULATION_TEMPLATE_NAME = 'business_process_regulation_template_v7.docx'
    private static final String TEMPLATE_FOLDER_NAME = 'Общие'

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

    //------------------------------------------------------------------------------------------------------------------
    // константы шаблона для таблиц
    //------------------------------------------------------------------------------------------------------------------
    private static final String ABBREVIATION_TEMPLATE_KEY = 'Сокращение'
    private static final String ABBREVIATION_VALUE_TEMPLATE_KEY = 'Значение сокращения'
    private static final String BUSINESS_PROCESS_LEVEL_TEMPLATE_KEY = 'Номер уровня БП'
    private static final String BUSINESS_PROCESS_CODE_TEMPLATE_KEY = 'Код БП'
    private static final String BUSINESS_PROCESS_NAME_TEMPLATE_KEY = 'Наименование БП'
    private static final String BUSINESS_PROCESS_PARAGRAPH_NUMBER_TEMPLATE_KEY = 'Номер  раздела для БП'
    private static final String EXTERNAL_BUSINESS_PROCESS_CODE_TEMPLATE_KEY = 'Код смежного БП'
    private static final String EXTERNAL_BUSINESS_PROCESS_NAME_TEMPLATE_KEY = 'Смежный БП'
    private static final String EXTERNAL_BUSINESS_PROCESS_INPUT_TEMPLATE_KEY = 'Вход из смежного БП'
    private static final String EXTERNAL_BUSINESS_PROCESS_OUTPUT_TEMPLATE_KEY = 'Выход в смежный БП'
    private static final String PROCESS_ROLE_TEMPLATE_KEY = 'Роль процесса'
    private static final String PROCESS_ROLE_POSITION_TEMPLATE_KEY = 'Должность для роли'
    private static final String PROCESS_ROLE_POSITION_ORGANIZATIONAL_UNIT_TEMPLATE_KEY = 'ОЕ для должности'
    private static final String PROCESS_DOCUMENT_COLLECTION_TEMPLATE_KEY = 'Набор документов'
    private static final String PROCESS_DOCUMENT_COLLECTION_CONTAINED_DOCUMENT_TEMPLATE_KEY = 'Документы набора'

    //------------------------------------------------------------------------------------------------------------------
    // константы шаблона для генерируемого раздела
    //------------------------------------------------------------------------------------------------------------------
    private static final String PROCESS_MODEL_TEMPLATE_KEY = 'Модель процесса'
    private static final String SCENARIO_CODE_TEMPLATE_KEY = 'Код сценария'
    private static final String SCENARIO_NAME_TEMPLATE_KEY = 'Сценарий'
    private static final String SCENARIO_REQUIREMENTS_TEMPLATE_KEY = 'Требования к сценарию'
    private static final String SCENARIO_MODEL_TEMPLATE_KEY = 'Модель сценария'
    private static final String PROCEDURE_NAME_TEMPLATE_KEY = 'Процедура'
    private static final String ROLE_NAME_TEMPLATE_KEY = 'Роль'
    private static final String PROCEDURE_CODE_TEMPLATE_KEY = 'Код процедуры'
    private static final String PROCEDURE_REQUIREMENTS_TEMPLATE_KEY = 'Требования к процедуре'
    private static final String PROCEDURE_MODEL_TEMPLATE_KEY = 'Модель процедуры'
    private static final String INPUT_DOCUMENT_EVENT_TEMPLATE_KEY = 'Входящий документ/событие'
    private static final String FUNCTION_TEMPLATE_KEY = 'Функция'
    private static final String OUTPUT_DOCUMENT_EVENT_TEMPLATE_KEY = 'Исходящий документ/событие'
    private static final String PERFORMER_TEMPLATE_KEY = 'Исполнитель'
    private static final String DURATION_TEMPLATE_KEY = 'Длительность'
    private static final String CHILD_FUNCTION_TEMPLATE_KEY = 'Условие'
    private static final String INFORMATION_SYSTEM_TEMPLATE_KEY = 'Информационная система'
    private static final String FUNCTION_REQUIREMENTS_TEMPLATE_KEY = 'Требования к функции'

    //------------------------------------------------------------------------------------------------------------------
    // константы id элементов
    //------------------------------------------------------------------------------------------------------------------
    private static final String ABBREVIATIONS_MODEL_ID = '0c25ad70-2733-11e6-05b7-db7cafd96ef7'
    private static final String ABBREVIATIONS_ROOT_OBJECT_ID = '0f7107e4-2733-11e6-05b7-db7cafd96ef7'
    private static final String FILE_REPOSITORY_ID = 'file-folder-root-id'
    private static final String FIRST_LEVEL_MODEL_ID = '1a8132f0-a43b-11e7-05b7-db7cafd96ef7'

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

    // TODO: [critical] определиться со списком исключаемых символов (интерфейсы, группировка интерфейсов и т.д.)
    private static final  List<String> EXCLUDED_FUNCTION_SYMBOL_IDS = [
            '07b15070-9b4e-4919-8ed0-9bae8764c7fa', // TODO: удалить? (интерфейс, созданный через редактор и генератор)
            '53a01270-95da-11ea-05b7-db7cafd96ef7', // TODO: удалить? (интерфейс СБП)
            '75f2e570-bdd3-11e5-05b7-db7cafd96ef7', // интерфейс смежного процесса
            'ST_PRCS_IF', // интерфейс процесса
            'fd841c20-cc37-11e6-05b7-db7cafd96ef7', // группировка интерфейсов
    ]
    // TODO: переименовать??? и уточнить по просто внешнему, а не смежному
    private static final String EXTERNAL_PROCESS_SYMBOL_ID = '75d9e6f0-4d1a-11e3-58a3-928422d47a25'
    private static final String NORMATIVE_DOCUMENT_SYMBOL_ID = '7096d320-cf42-11e2-69e4-ac8112d1b401'
    // TODO: уточнить по другим типам символа сценария
    private static final String SCENARIO_SYMBOL_ID = 'ST_SCENARIO'
    private static final String STATUS_SYMBOL_ID = 'd6e8a7b0-7ce6-11e2-3463-e4115bf4fdb9'

    //------------------------------------------------------------------------------------------------------------------
    // константы для отладки при разработке
    //------------------------------------------------------------------------------------------------------------------
    private static final boolean DEBUG = true
    private static final String TEMPLATE_LOCAL_PATH = 'C:\\Users\\vikto\\IdeaProjects\\BusinessProcessRegulationRemake\\examples'


    //------------------------------------------------------------------------------------------------------------------
    // основной код
    //------------------------------------------------------------------------------------------------------------------
    private static final int CM_1_OFFSET = 567 // 1 сантиметр (отступ в документе)

    private static Map<String, String> fullAbbreviations = new TreeMap<>()
    private static Pattern abbreviationsPattern = null
    private static Map<String, String> foundedAbbreviations = new TreeMap<>()

    CustomScriptContext context
    private TreeRepository treeRepository

    private static int detailLevel = 3
    private static String docVersion = ''
    private static String docDate = ''
    private static String currentYear = LocalDate.now().getYear().toString()

    enum SubprocessOwnerType {
        ORGANIZATIONAL_UNIT,
        GROUP,
    }

    private static final Map<String, SubprocessOwnerType> subprocessOwnerTypeMap = Map.of(
            ORGANIZATIONAL_UNIT_OBJECT_TYPE_ID, SubprocessOwnerType.ORGANIZATIONAL_UNIT,
            GROUP_OBJECT_TYPE_ID, SubprocessOwnerType.GROUP,
    )

    private class CommonObjectInfo {
        ObjectElement object
        String name

        CommonObjectInfo(ObjectElement object) {
            this.object = object
            this.name = getName(object.getObjectDefinition())
        }

        CommonObjectInfo(Model model) {
            this.object = null
            this.name = getName(model)
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
            // TODO: [critical] логика получения имени, требований, кода для одиночных сценариев
            this.function = new CommonObjectInfo(model)
            this.code = getAttributeValue(model, DATA_ELEMENT_CODE_ATTR_ID)
            this.requirements = getAttributeValue(model, DESCRIPTION_DEFINITION_ATTR_ID)
        }
    }

    private class PositionInfo {
        CommonObjectInfo position
        // TODO: уточнить, одна ли ОЕ?
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
        // TODO: уточнить, может ли быть несколько?
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

        void findContainedDocuments () {
            ObjectElement collectionObjectOnModel = model.findObjectInstances(collection.document.object.getObjectDefinition())
                    .stream()
                    .findFirst()
                    .orElse(null)

            if (collectionObjectOnModel == null) {
                return
            }

            // TODO: обсудить логику определения состава коллекции (пример с отсутсвием связей для части доков на модели; c неправильным направлением связи)
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
        String fileName
        FileNodeDTO content = null

        FileInfo(String fileName, FileNodeDTO content) {
            this.fileName = fileName
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
                }
                else {
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
                    .findAll { ObjectElement oE -> oE.getSymbolId() == SCENARIO_SYMBOL_ID }
                    .unique(Comparator.comparing { ObjectElement oE -> oE.getId() })
                    .sort { ObjectElement oE1, ObjectElement oE2 -> ModelUtils.getElementsCoordinatesComparator().compare(oE1, oE2) }

            scenarioObjects.each { ObjectElement scenarioObject ->
                Model scenarioModel = getEPCModel(scenarioObject)

                if (scenarioModel == null) {
                    // TODO: [critical] что делать, если у какого-либо сценария нет декомпозиции?
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
                    // TODO: [critical] что делать, если выбран режим до 4 уровня, а у какого-либо 3 лвла нет декомпозиции на 4?
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
                List<String> procedureBusinessRoleNames = procedure.businessRoles.collect {BusinessRoleInfo businessRole -> businessRole.businessRole.name}
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

        void findNormativeDocuments () {
            List<ObjectElement> normativeDocumentObjects = model.getObjects()
                    .findAll { ObjectElement oE -> oE.getSymbolId() == NORMATIVE_DOCUMENT_SYMBOL_ID }
                    .unique(Comparator.comparing { ObjectElement oE -> oE.getObjectDefinitionId() })
            normativeDocuments = normativeDocumentObjects.collect { ObjectElement normativeDocumentObject -> new NormativeDocumentInfo(normativeDocumentObject) }
        }

        void findDocumentCollections () {
            List<ObjectElement> documentObjects = model.getObjects()
                    .findAll { ObjectElement oE -> oE.getObjectDefinition().getObjectTypeId() in DOCUMENT_OBJECT_TYPE_IDS }
                    .unique(Comparator.comparing { ObjectElement oE -> oE.getObjectDefinitionId() })
                    .sort { ObjectElement oE1, ObjectElement oE2 -> ModelUtils.getElementsCoordinatesComparator().compare(oE1, oE2) }

            documentObjects.each { ObjectElement documentObject ->
                Model documentCollectionModel = findDocumentCollectionModel(documentObject)

                // TODO: обсудить логику определения набора документов (пример ошибки с типом символа)
                if (documentCollectionModel) {
                    documentCollections.add(new DocumentCollectionInfo(documentObject, documentCollectionModel))
                }
            }
        }

        static Model findDocumentCollectionModel(ObjectElement documentCollectionObject) {
            List <Model> documentCollectionObjectModels = documentCollectionObject.getDecompositions()
                    .findAll { TreeNode tN -> tN.isModel() } as List <Model>
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

            if (!epcFunction.inputDocuments) {
                epcFunction.findInputEvents()
            }

            epcFunction.findOutputDocuments()

            if (!epcFunction.outputDocuments) {
                epcFunction.findOutputEvents()
            }

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
                    .unique(Comparator.comparing { ObjectElement oE -> oE.getObjectDefinitionId() })
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
                    .unique(Comparator.comparing { ObjectElement oE -> oE.getObjectDefinitionId() })
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
            informationSystems = informationSystemObjects.collect { ObjectElement informationSystemObject -> new CommonObjectInfo(informationSystemObject) }
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
                }
                else {
                    EPCFunctionDescription childEPCFunction = epcFunctions
                            .find { EPCFunctionDescription epcFunction -> epcFunction.function.function.object.getId() == childFunctionObject.getId() }
                    childEPCFunctions.add(childEPCFunction)
                }
            }
        }

        private List<ObjectElement> findOutputEventObjects (ObjectElement functionObject) {
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

        FileNodeDTO content = null

        BusinessProcessRegulationDocument(String fileName, SubprocessDescription subprocessDescription, XWPFDocument template, int detailLevel) {
            this.fileName = fileName
            this.subprocessDescription = subprocessDescription
            this.document = template
            this.detailLevel = detailLevel
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
            String docVersionWithDateTemplateValue = "${docVersion} от ${docDate}"
            return Map.of(
                    PROCESS_NAME_UPPER_CASE_TEMPLATE_KEY, subprocessDescription.subprocess.function.name.toUpperCase(),
                    PROCESS_CODE_TEMPLATE_KEY, subprocessDescription.subprocess.code,
                    DOC_VERSION_WITH_DATE_TEMPLATE_KEY, docVersionWithDateTemplateValue,
                    DOC_YEAR_TEMPLATE_KEY, currentYear,
                    DOC_VERSION_TEMPLATE_KEY, docVersion,
                    DOC_DATE_TEMPLATE_KEY, docDate,
                    PROCESS_NAME_TEMPLATE_KEY, subprocessDescription.subprocess.function.name,
                    FIRST_LEVEL_PROCESS_NAME_TEMPLATE_KEY, subprocessDescription.parentProcess.function.name,
                    PROCESS_REQUIREMENTS_TEMPLATE_KEY, subprocessDescription.subprocess.requirements,
            )
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
                replacement += owner.owner.name ? "«${owner.owner.name}»" : "<${PROCESS_OWNER_TEMPLATE_KEY}>"

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
            boolean replacementWas = false
            for (normativeDocument in subprocessDescription.completedNormativeDocuments) {
                String replacement = normativeDocument.requisites ? normativeDocument.requisites : pattern
                boolean replacementFlag = replaceInCopyParagraph(document, pattern, replacement)
                replacementWas = replacementFlag ? replacementFlag : replacementWas
            }

            if (replacementWas) {
                removeParagraphByText(document, pattern)
            }
        }

        private void fillProcessOwners() {
            String pattern = "<${PROCESS_OWNER_TEMPLATE_KEY}>"
            String placeholderCopy = ", ${pattern}"
            for (owner in subprocessDescription.owners) {
                String replacement = owner.owner.name ? "«${owner.owner.name}»" : "<${PROCESS_OWNER_TEMPLATE_KEY}>"

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

            int counter = 0
            boolean replacementWas = false
            for (goal in subprocessDescription.goals) {
                counter+=1
                String replacement = goal.name ? goal.name : pattern

                if (counter == goalsCount) {
                    pattern += ';'
                    replacement += '.'
                }

                boolean replacementFlag = replaceInCopyParagraph(document, pattern, replacement)
                replacementWas = replacementFlag ? replacementFlag : replacementWas
            }

            if (replacementWas) {
                removeParagraphByText(document, pattern)
            }
        }

        void fillTables() {
            fillAbbreviations()
            fillBusinessProcessHierarchy()
            fillExternalBusinessProcesses()
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

            if (foundedAbbreviations) {
                table.removeRow(1)
            }
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
            String nameReplacement = "\\L${(level.toInteger() - 1).toString()}${code} ${name}"

            XWPFTableRow newTableRow = copyTableRow(table.getRows().get(1), table)
            replaceParagraphText(newTableRow.getTableCells().get(0).getParagraphs().get(0), levelPattern, level)
            replaceParagraphText(newTableRow.getTableCells().get(1).getParagraphs().get(0), namePattern, nameReplacement)
        }

        private void fillExternalBusinessProcesses() {
            XWPFTable table = findTableByHeaders(document, EXTERNAL_BUSINESS_PROCESS_TABLE_HEADERS)

            if (table.getRows().size() != 3) {
                return
            }

            for (externalProcessDescription in subprocessDescription.completedExternalProcessesWithInputFlows) {
                fillExternalBusinessProcess(table, externalProcessDescription, EXTERNAL_BUSINESS_PROCESS_INPUT_TEMPLATE_KEY, 1)
            }

            for (externalProcessDescription in subprocessDescription.completedExternalProcessesWithOutputFlows) {
                fillExternalBusinessProcess(table, externalProcessDescription, EXTERNAL_BUSINESS_PROCESS_OUTPUT_TEMPLATE_KEY, 2)
            }

            table.removeRow(1)
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
            for (flowName in flowNames) {
                String flowReplacement = "${flowName};"
                replaceInCopyParagraph(newTableRow.getTableCells().get(flowColumnNumber), flowPattern, flowReplacement)
            }

            newTableRow.getTableCells().get(flowColumnNumber).removeParagraph(newTableRow.getTableCells().get(flowColumnNumber).getParagraphs().size() - 1)
        }

        private static XWPFTable findTableByHeaders(XWPFDocument document, List<String> tableHeaders) {
            XWPFTable foundedTable = null
            for (table in document.getTables()) {
                if (table.getRows().size() == 0) {
                    continue
                }

                if (table.getRows().get(0).getTableCells().size() != tableHeaders.size()) {
                    continue
                }

                boolean tableFound = true
                for (int columnNumber = 0; columnNumber < tableHeaders.size(); columnNumber++) {
                    String header = tableHeaders[columnNumber]
                    XWPFTableCell cell = table.getRows().get(0).getTableCells().get(columnNumber)

                    if (cell.getParagraphs().size() == 0) {
                        tableFound = false
                        break
                    }

                    XWPFParagraph paragraph = cell.getParagraphs().get(0)
                    if (paragraph.getText() != header) {
                        tableFound = false
                        break
                    }
                }

                if (tableFound) {
                    foundedTable = table
                    break
                }
            }
            return foundedTable
        }

        private static XWPFTableRow copyTableRow(XWPFTableRow sourceRow, XWPFTable table) {
            XWPFTableRow newRow = table.createRow()
            newRow.getCtRow().setTrPr(sourceRow.getCtRow().getTrPr())

            for (int cellNumber = 0; cellNumber < sourceRow.getTableCells().size(); cellNumber++) {
                XWPFTableCell targetCell = newRow.getTableCells().get(cellNumber)
                XWPFTableCell sourceCell = sourceRow.getTableCells().get(cellNumber)
                targetCell.getCTTc().setTcPr(sourceCell.getCTTc().getTcPr())

                for (int paragraphNumber = 0; paragraphNumber < sourceCell.getParagraphs().size(); paragraphNumber++) {
                    XWPFParagraph sourceParagraph = sourceCell.getParagraphs().get(paragraphNumber)
                    XWPFParagraph targetParagraph = targetCell.getParagraphs().get(paragraphNumber)
                    copyParagraph(sourceParagraph, targetParagraph)
                }
            }
            return newRow
        }

        private void replaceParagraphsText(String pattern, String replacement) {
            for (paragraph in document.getParagraphs()) {
                replaceParagraphText(paragraph, pattern, replacement)
            }
        }

        private void replaceHeadersText(String pattern, String replacement) {
            for (header in document.getHeaderList()) {
                for (headerParagraph in header.getParagraphs()) {
                    replaceParagraphText(headerParagraph, pattern, replacement)
                }
            }
        }

        private static boolean replaceInCopyParagraph(IBody body, String pattern, String replacement) {
            if (pattern == replacement) {
                return false
            }

            List<XWPFParagraph> paragraphs = findParagraphsByText(body, pattern)
            paragraphs.each { XWPFParagraph paragraph ->
                XWPFParagraph newParagraph = addParagraph(body, paragraph)
                replaceParagraphText(newParagraph, pattern, replacement)
            }
            return true
        }

        private static List<XWPFParagraph> findParagraphsByText(IBody body, String text) {
            return body.getParagraphs()
                    .findAll { XWPFParagraph paragraph -> paragraph.getText().contains(text) }
        }

        private static XWPFParagraph addParagraph(IBody body, XWPFParagraph paragraph) {
            XmlCursor cursor = paragraph.getCTP().newCursor()
            XWPFParagraph newParagraph = body.insertNewParagraph(cursor)
            copyParagraph(paragraph, newParagraph)
            return newParagraph
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

        private static void removeParagraphByText(XWPFDocument document, String pattern) {
            List<XWPFParagraph> paragraphs = findParagraphsByText(document, pattern)
            paragraphs.each { XWPFParagraph paragraph ->
                removeParagraph(document, paragraph)
            }
        }

        private static void removeParagraph(XWPFDocument document, XWPFParagraph paragraph) {
            int position = document.getPosOfParagraph(paragraph)
            document.removeBodyElement(position)
        }

        private static void replaceParagraphText(XWPFParagraph paragraph, String pattern, String replacement) {
            if (paragraph.getText().contains(pattern)) {
                String newText = paragraph.getText().replace(pattern, replacement)
                addParagraphText(paragraph, newText)
            }
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

            run.setText(text, 0)
            paragraph.addRun(run)
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
            if (DEBUG) {
                FileOutputStream file = new FileOutputStream("${TEMPLATE_LOCAL_PATH}\\${fileName}.${DOCX_FORMAT}")
                document.write(file)
                file.close()
            }

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

    private static String getName(TreeNode treeNode) {
        Node node = treeNode._getNode()
        String name = node.getName()
        name = name ? trimStringValue(name) : ''
        if (name) {
            findAbbreviations(name)
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
            resultFileName = files[0].fileName
            format = DOCX_FORMAT
            result = files[0].content
        }

        if (files.size() > 1) {
            resultFileName = ZIP_RESULT_FILE_NAME_FIRST_PART + new SimpleDateFormat('yyyyMMdd HHmmss').format(new Date()).replace(' ', '_')
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
        treeRepository = context.createTreeRepository(true)
        parseParameters()
        initAbbreviations()
    }

    private void parseParameters() {
        if (DEBUG) {
            detailLevel = 4
            docVersion = '01'
            docDate = '01.01.2025'
            return
        }

        String deep = ParamUtils.parse(context.findParameter(DETAIL_LEVEL_PARAM_NAME)) as String
        detailLevel = Integer.parseInt(deep.replaceAll('[^0-9]+', ''))

        docVersion = ParamUtils.parse(context.findParameter(DOC_VERSION_PARAM_NAME)) as String

        Timestamp approvalDate = ParamUtils.parse(context.findParameter(DOC_DATE_PARAM_NAME)) as Timestamp
        docDate = approvalDate.format('dd.MM.yyyy')
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
        //noinspection RegExpUnnecessaryNonCapturingGroup
        abbreviationsPattern = Pattern.compile("\\b(?:(?:${String.join(')|(?:', abbreviationNames)}))\\b")
    }

    private List<ObjectElement> getSubprocessObjects() {
        List<ObjectElement> subprocessObjects = []
        if (!context.elementsIdsList().isEmpty()){
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
        List<SubprocessDescription> subprocessDescriptions = subprocessObjects.collect{ ObjectElement subprocessObject -> new SubprocessDescription(subprocessObject, detailLevel) }
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
        subprocess.defineNormativeDocuments()
        subprocess.completeNormativeDocuments()
        subprocess.defineDocumentCollections()
        subprocess.completeDocumentCollections()
        subprocess.analyzeEPCModels()
    }

    private List<FileInfo> createBPRegulations(List<SubprocessDescription> subprocessDescriptions) {
        List<FileInfo> files = []
        subprocessDescriptions.forEach { SubprocessDescription subprocessDescription ->
            String fileName = DOCX_RESULT_FILE_NAME_FIRST_PART + " '${subprocessDescription.subprocess.function.name}' " + new SimpleDateFormat('yyyyMMdd HHmmss').format(new Date()).replace(' ', '_')
            BusinessProcessRegulationDocument document = getBPRegulationDocument(fileName, subprocessDescription)
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
                return new XWPFDocument(fileInputStream)
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
        return new XWPFDocument(new ByteArrayInputStream(file))
    }

    private byte[] createZipFileContent(List<FileInfo> files) {
        ByteArrayOutputStream byteArrayOutputStream = new ByteArrayOutputStream()
        ZipOutputStream zipOutputStream = new ZipOutputStream(byteArrayOutputStream)

        files.each { FileInfo file ->
            ZipEntry zipEntry = new ZipEntry(file.fileName + '.docx')
            zipOutputStream.putNextEntry(zipEntry)
            zipOutputStream.write(file.content.file.bytes, 0, file.content.file.bytes.length)
            zipOutputStream.closeEntry()
        }

        zipOutputStream.close()
        byteArrayOutputStream.close()

        return byteArrayOutputStream.toByteArray()
    }
}
