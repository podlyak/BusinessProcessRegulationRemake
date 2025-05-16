import groovy.util.logging.Slf4j
import ru.nextconsulting.bpm.dto.NodeId
import ru.nextconsulting.bpm.repository.business.AttributeValue
import ru.nextconsulting.bpm.repository.structure.Node
import ru.nextconsulting.bpm.repository.structure.ObjectDefinitionNode
import ru.nextconsulting.bpm.repository.structure.ScriptParameter
import ru.nextconsulting.bpm.repository.structure.SilaScriptParamType
import ru.nextconsulting.bpm.script.repository.TreeRepository
import ru.nextconsulting.bpm.script.tree.elements.ObjectElement
import ru.nextconsulting.bpm.script.tree.node.Model
import ru.nextconsulting.bpm.script.tree.node.ObjectDefinition
import ru.nextconsulting.bpm.script.tree.node.TreeNode
import ru.nextconsulting.bpm.script.utils.ModelUtils
import ru.nextconsulting.bpm.scriptengine.context.ContextParameters
import ru.nextconsulting.bpm.scriptengine.context.CustomScriptContext
import ru.nextconsulting.bpm.scriptengine.exception.SilaScriptException
import ru.nextconsulting.bpm.scriptengine.script.GroovyScript
import ru.nextconsulting.bpm.scriptengine.util.ParamUtils
import ru.nextconsulting.bpm.scriptengine.util.SilaScriptParameter
import ru.nextconsulting.bpm.scriptengine.util.SilaScriptParameters
import ru.nextconsulting.bpm.utils.JsonConverter

import java.sql.Timestamp
import java.time.LocalDate
import java.util.regex.Matcher
import java.util.regex.Pattern

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

    private static final String DETAIL_LEVEL_PARAM_NAME = 'Глубина детализации регламента'
    private static final String DOC_VERSION_PARAM_NAME = 'Номер версии регламента'
    private static final String DOC_DATE_PARAM_NAME = 'Дата утверждения регламента'

    private static final String ABBREVIATIONS_MODEL_ID = '0c25ad70-2733-11e6-05b7-db7cafd96ef7'
    private static final String ABBREVIATIONS_ROOT_OBJECT_ID = '0f7107e4-2733-11e6-05b7-db7cafd96ef7'
    private static final String FIRST_LEVEL_MODEL_ID = '1a8132f0-a43b-11e7-05b7-db7cafd96ef7'

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

    private static final String BUSINESS_ROLE_OBJECT_TYPE_ID = 'OT_PERS_TYPE'
    private static final String CLUSTER_DATA_MODEL_OBJECT_TYPE_ID = 'OT_CLST'
    private static final String FLOW_OBJECT_TYPE_ID = 'OT_TECH_TRM'
    private static final String FUNCTION_OBJECT_TYPE_ID = 'OT_FUNC'
    private static final String GOAL_OBJECT_TYPE_ID = 'OT_OBJECTIVE'
    private static final String GROUP_OBJECT_TYPE_ID = 'OT_GRP'
    private static final String INFORMATION_CARRIER_OBJECT_TYPE_ID = 'OT_INFO_CARR'
    private static final String ORGANIZATIONAL_UNIT_OBJECT_TYPE_ID = 'OT_ORG_UNIT'

    private static final List<String> DOCUMENT_COLLECTION_OBJECT_TYPE_IDS = [
            CLUSTER_DATA_MODEL_OBJECT_TYPE_ID,
            INFORMATION_CARRIER_OBJECT_TYPE_ID,
    ]

    private static final List<String> ABBREVIATIONS_EDGE_TYPE_IDS = [
            'CT_IS_IN_RELSHP_TO',
            'CT_IS_IN_RELSHP_TO_1',
            'CT_HAS_REL_WITH',
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
    private static final String INPUT_FLOW_W_SUBPROCESS_EDGE_TYPE_ID = 'CT_IS_INP_FOR'
    private static final String LEADERSHIP_POSITION_W_OWNER_EDGE_TYPE_ID = 'CT_IS_DISC_SUPER'
    private static final String ORGANIZATIONAL_UNIT_W_POSITION_EDGE_TYPE_ID = 'CT_IS_CRT_BY'
    private static final String OUTPUT_FLOW_W_CUSTOMER_EDGE_TYPE_ID = 'CT_IS_INP_FOR'
    private static final List<String> OWNER_W_SUBPROCESS_EDGE_TYPE_IDS = [
            'CT_EXEC_1',
            'CT_EXEC_2',
    ]
    private static final String POSITION_W_BUSINESS_ROLE_EDGE_TYPE_ID = 'CT_EXEC_5'
    private static final String SUBPROCESS_W_OUTPUT_FLOW_EDGE_TYPE_ID = 'CT_HAS_OUT'
    private static final String SUPPLIER_W_INPUT_FLOW_EDGE_TYPE_ID = 'CT_HAS_OUT'

    private static final String DATA_ELEMENT_CODE_ATTR_ID = '46e148b0-b96d-11e3-05b7-db7cafd96ef7'
    private static final String DESCRIPTION_DEFINITION_ATTR_ID = 'AT_DESC'
    private static final String FULL_NAME_ATTR_ID = 'AT_NAME_FULL'

    // TODO: переименовать??? и уточнить по просто внешнему, а не смежному
    private static final String EXTERNAL_PROCESS_SYMBOL_ID = '75d9e6f0-4d1a-11e3-58a3-928422d47a25'
    private static final String NORMATIVE_DOCUMENT_SYMBOL_ID = '7096d320-cf42-11e2-69e4-ac8112d1b401'
    // TODO: уточнить по другим типам символа сценария
    private static final String SCENARIO_SYMBOL_ID = 'ST_SCENARIO'

    private static Map<String, String> fullAbbreviations = new TreeMap<>()
    private static Pattern abbreviationsPattern = null
    private static Map<String, String> foundedAbbreviations = new TreeMap<>()

    CustomScriptContext context
    private TreeRepository treeRepository

    private static int detailLevel = 3
    private static String docVersion = ''
    private static String docDate = ''
    private static String currentYear = LocalDate.now().getYear().toString()

    private static boolean debug = true

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
            // TODO: логика получения имени, требований, кода для одиночных сценариев
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
                        .findAll {it.getEdgeTypeId() == ORGANIZATIONAL_UNIT_W_POSITION_EDGE_TYPE_ID}
                        .collect {it.getSource() as ObjectElement}
                        .unique(Comparator.comparing { ObjectElement o -> o.getId() })
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
                        .findAll {it.getEdgeTypeId() == POSITION_W_BUSINESS_ROLE_EDGE_TYPE_ID}
                        .collect {it.getSource() as ObjectElement}
                        .unique(Comparator.comparing { ObjectElement o -> o.getId() })
                positions.addAll(positionObjects.collect {new PositionInfo(it)})
            }
        }
    }

    private class DocumentInfo {
        CommonObjectInfo document
        String type

        DocumentInfo(ObjectElement document) {
            this.document = new CommonObjectInfo(document)
            this.type = document.getSymbol().name
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

        private void findContainedDocuments () {
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
                        .findAll {it.getEdgeTypeId() in CLUSTER_GROUP_W_CLUSTER_DATA_MODEL_EDGE_TYPE_IDS}
                        .collect {it.getTarget() as ObjectElement}
                        .findAll {it.getObjectDefinition().getObjectTypeId() == CLUSTER_DATA_MODEL_OBJECT_TYPE_ID}
                        .unique(Comparator.comparing { ObjectElement o -> o.getObjectDefinitionId() })
                        .sort {o1, o2 -> ModelUtils.getElementsCoordinatesComparator().compare(o1, o2)}
                )
            }

            if (modelTypeId == INFORMATION_CARRIER_MODEL_TYPE_ID) {
                containedDocumentObjects.addAll(collectionObjectOnModel.getExitEdges()
                        .findAll {it.getEdgeTypeId() in DOCUMENT_COLLECTION_W_DOCUMENT_EDGE_TYPE_IDS}
                        .collect {it.getTarget() as ObjectElement}
                        .findAll {it.getObjectDefinition().getObjectTypeId() == INFORMATION_CARRIER_OBJECT_TYPE_ID}
                        .unique(Comparator.comparing { ObjectElement o -> o.getObjectDefinitionId() })
                        .sort {o1, o2 -> ModelUtils.getElementsCoordinatesComparator().compare(o1, o2)}
                )
            }

            containedDocuments = containedDocumentObjects.collect {new DocumentInfo(it)}
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
        List<SubprocessOwnerDescription> owners = []
        List<CommonObjectInfo> goals = []
        List<InputFlowDescription> externalProcessInputFlowDescriptions = []
        List<OutputFlowDescription> externalProcessOutputFlowDescriptions = []
        Model processSelectionModel = null
        List<ScenarioDescription> scenarios = []

        List<ExternalProcessDescription> externalProcessesWithInputFlows = []
        List<ExternalProcessDescription> externalProcessesWithOutputFlows = []
        List<EPCDescription> analyzedEPC = []
        List<DocumentCollectionInfo> fullDocumentCollections = []

        SubprocessDescription(ObjectElement subprocess, int detailLevel) {
            this.subprocess = new CommonFunctionInfo(subprocess)
            this.detailLevel = detailLevel
        }

        private void defineParentProcess() {
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

        private void findOwners() {
            List<ObjectElement> ownerObjects = subprocess.function.object.getEnterEdges()
                    .findAll {it.getEdgeTypeId() in OWNER_W_SUBPROCESS_EDGE_TYPE_IDS}
                    .collect {it.getSource() as ObjectElement}
                    .unique(Comparator.comparing { ObjectElement o -> o.getId() })
                    .sort {o1, o2 -> ModelUtils.getElementsCoordinatesComparator().compare(o1, o2)}
            owners = ownerObjects.collect {
                new SubprocessOwnerDescription(it, subprocessOwnerTypeMap.get(it.getObjectDefinition().getObjectTypeId()))
            }
        }

        private void defineGoals() {
            Model functionAllocationModel = subprocess.function.object.getObjectDefinition()
                    .getDecompositions(FUNCTION_ALLOCATION_MODEL_TYPE_ID)
                    .stream()
                    .findFirst()
                    .orElse(null)

            if (functionAllocationModel == null) {
                return
            }

            List<ObjectElement> goalObjects = functionAllocationModel.findObjectsByType(GOAL_OBJECT_TYPE_ID)
            goals = goalObjects.collect {new CommonObjectInfo(it)}
        }

        private void findExternalProcessInputFlows() {
            List<ObjectElement> allFlowObjects = subprocess.function.object.model.findObjectsByType(FLOW_OBJECT_TYPE_ID)

            List<ObjectElement> inputFlowObjects = subprocess.function.object.getEnterEdges()
                    .findAll {it.getEdgeTypeId() == INPUT_FLOW_W_SUBPROCESS_EDGE_TYPE_ID}
                    .collect {it.getSource() as ObjectElement}
                    .findAll {it.getObjectDefinition().getObjectTypeId() == FLOW_OBJECT_TYPE_ID}
                    .unique(Comparator.comparing { ObjectElement o -> o.getId() })
                    .sort {o1, o2 -> ModelUtils.getElementsCoordinatesComparator().compare(o1, o2)}

            inputFlowObjects.each {ObjectElement currentFlowObject ->
                List<ObjectElement> externalSupplierObjects = currentFlowObject.getEnterEdges()
                        .findAll {it.getEdgeTypeId() == SUPPLIER_W_INPUT_FLOW_EDGE_TYPE_ID}
                        .collect {it.getSource() as ObjectElement}
                        .findAll {it.getSymbolId() == EXTERNAL_PROCESS_SYMBOL_ID}
                        .unique(Comparator.comparing { ObjectElement o -> o.getId() })
                        .sort {o1, o2 -> ModelUtils.getElementsCoordinatesComparator().compare(o1, o2)}

                List<ObjectElement> additionalExternalSupplierObjects = findAdditionalExternalSupplierObjects(currentFlowObject, allFlowObjects)
                externalSupplierObjects.addAll(additionalExternalSupplierObjects)
                externalSupplierObjects = externalSupplierObjects
                        .unique(Comparator.comparing { ObjectElement o -> o.getObjectDefinitionId() })
                        .sort {o1, o2 -> ModelUtils.getElementsCoordinatesComparator().compare(o1, o2)}

                if (externalSupplierObjects) {
                    externalProcessInputFlowDescriptions.add(new InputFlowDescription(currentFlowObject, externalSupplierObjects))
                }
            }
        }

        private List<ObjectElement> findAdditionalExternalSupplierObjects(ObjectElement currentFlowObject, List<ObjectElement> allFlowObjects) {
            String currentFlowObjectDefinitionId = currentFlowObject.getObjectDefinitionId()
            List<ObjectElement> currentFlowObjects = allFlowObjects
                    .findAll {it.getObjectDefinitionId() == currentFlowObjectDefinitionId}

            List<ObjectElement> additionalExternalSupplierObjects = []
            for (flowObject in currentFlowObjects) {
                if (flowObject.getId() == currentFlowObject.getId()) {
                    continue
                }

                List<ObjectElement> foundedObjects = flowObject.getEnterEdges()
                        .findAll {it.getEdgeTypeId() == SUBPROCESS_W_OUTPUT_FLOW_EDGE_TYPE_ID}
                        .collect {it.getSource() as ObjectElement}
                        .findAll {it.getSymbolId() == EXTERNAL_PROCESS_SYMBOL_ID}
                        .unique(Comparator.comparing { ObjectElement o -> o.getId() })
                        .sort {o1, o2 -> ModelUtils.getElementsCoordinatesComparator().compare(o1, o2)}
                additionalExternalSupplierObjects.addAll(foundedObjects)
            }
            return additionalExternalSupplierObjects
        }

        private void findExternalProcessOutputFlows() {
            List<ObjectElement> outputFlowObjects = subprocess.function.object.getExitEdges()
                    .findAll {it.getEdgeTypeId() == SUBPROCESS_W_OUTPUT_FLOW_EDGE_TYPE_ID}
                    .collect {it.getTarget() as ObjectElement}
                    .findAll {it.getObjectDefinition().getObjectTypeId() == FLOW_OBJECT_TYPE_ID}
                    .unique(Comparator.comparing { ObjectElement o -> o.getId() })
                    .sort {o1, o2 -> ModelUtils.getElementsCoordinatesComparator().compare(o1, o2)}

            outputFlowObjects.each {ObjectElement currentFlowObject ->
                List<ObjectElement> externalCustomerObjects = currentFlowObject.getExitEdges()
                        .findAll {it.getEdgeTypeId() == OUTPUT_FLOW_W_CUSTOMER_EDGE_TYPE_ID}
                        .collect {it.getTarget() as ObjectElement}
                        .findAll {it.getSymbolId() == EXTERNAL_PROCESS_SYMBOL_ID}
                        .unique(Comparator.comparing { ObjectElement o -> o.getId() })
                        .sort {o1, o2 -> ModelUtils.getElementsCoordinatesComparator().compare(o1, o2)}

                if (externalCustomerObjects) {
                    externalProcessOutputFlowDescriptions.add(new OutputFlowDescription(currentFlowObject, externalCustomerObjects))
                }
            }
        }

        private void buildExternalProcessesWithInputFlows() {
            externalProcessInputFlowDescriptions.each {InputFlowDescription inputFlowDescription ->
                addExternalProcessesWithFlow(inputFlowDescription.inputFlow, inputFlowDescription.suppliers, externalProcessesWithInputFlows)
            }
        }

        private void buildExternalProcessesWithOutputFlows() {
            externalProcessOutputFlowDescriptions.each {OutputFlowDescription outputFlowDescription ->
                addExternalProcessesWithFlow(outputFlowDescription.outputFlow, outputFlowDescription.customers, externalProcessesWithOutputFlows)
            }
        }

        private void addExternalProcessesWithFlow(CommonObjectInfo flow, List<CommonFunctionInfo> externalProcesses, List<ExternalProcessDescription> externalProcessesWithFlows) {
            for (process in externalProcesses) {
                List<String> addedProcessObjectDefinitionIds = externalProcessesWithFlows.collect {it.externalProcess.function.object.getObjectDefinitionId()}
                String processObjectDefinitionId = process.function.object.getObjectDefinitionId()

                if (processObjectDefinitionId in addedProcessObjectDefinitionIds) {
                    ExternalProcessDescription processDescription = externalProcessesWithFlows
                            .find {it.externalProcess.function.object.getObjectDefinitionId() == processObjectDefinitionId}

                    List<String> addedFlowObjectDefinitionIds = processDescription.flows.collect {it.object.getObjectDefinitionId()}
                    if (flow.object.getObjectDefinitionId() in addedFlowObjectDefinitionIds) {
                        continue
                    }

                    processDescription.flows.add(flow)
                }
                else {
                    externalProcessesWithFlows.add(new ExternalProcessDescription(process, [flow]))
                }
            }
        }

        private void defineProcessSelectionModel() {
            processSelectionModel = subprocess.function.object.getObjectDefinition()
                    .getDecompositions(PROCESS_SELECTION_MODEL_TYPE_ID)
                    .stream()
                    .findFirst()
                    .orElse(null)
        }

        private void defineScenarios() {
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
                    .findAll {it.getSymbolId() == SCENARIO_SYMBOL_ID}
                    .unique(Comparator.comparing { ObjectElement o -> o.getId() })
                    .sort {o1, o2 -> ModelUtils.getElementsCoordinatesComparator().compare(o1, o2)}

            scenarioObjects.each {ObjectElement scenarioObject ->
                Model scenarioModel = getEPCModel(scenarioObject)

                if (scenarioModel == null) {
                    // TODO: что делать, если у какого-либо сценария нет декомпозиции?
                    return
                }

                scenarios.add(new ScenarioDescription(scenarioModel, scenarioObject))
            }
        }

        private void defineProcedures() {
            scenarios.each {it.defineProcedures()}
        }

        private void defineBusinessRoles() {
            scenarios.each {it.defineBusinessRoles()}
        }

        private void identifyAnalyzedEPC() {
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

        private void defineNormativeDocuments() {
            analyzedEPC.each {it.findNormativeDocuments()}
        }

        private void defineDocumentCollections() {
            analyzedEPC.each {it.findDocumentCollections()}
        }

        private void buildFullDocumentCollections() {
            for (epcDescription in analyzedEPC) {
                for (documentCollection in epcDescription.documentCollections) {
                    fullDocumentCollections.add(documentCollection)
                }
            }

            fullDocumentCollections = fullDocumentCollections
                    .unique(Comparator.comparing { DocumentCollectionInfo o -> o.collection.document.object.getObjectDefinitionId() })

            List<DocumentCollectionInfo> foundedDocumentCollections = fullDocumentCollections
            while (foundedDocumentCollections) {
                List<DocumentCollectionInfo> unparsedDocumentCollections = foundedDocumentCollections
                unparsedDocumentCollections.each {it.findContainedDocuments()}
                foundedDocumentCollections = parseDocumentCollections(unparsedDocumentCollections)
                fullDocumentCollections.addAll(foundedDocumentCollections)
            }
        }

        private List<DocumentCollectionInfo> parseDocumentCollections(List<DocumentCollectionInfo> unparsedDocumentCollections) {
            List<String> fullDocumentCollectionsObjectDefinitionIds = fullDocumentCollections.collect { it.collection.document.object.getObjectDefinitionId()}
            List<DocumentCollectionInfo> foundedDocumentCollections = []
            for (unparsedDocumentCollection in unparsedDocumentCollections) {
                for (containedDocument in unparsedDocumentCollection.containedDocuments) {
                    Model containedDocumentModel = EPCDescription.findDocumentCollectionModel(containedDocument.document.object)
                    boolean containedDocumentAlreadyInFullCollections = containedDocument.document.object.getObjectDefinitionId() in fullDocumentCollectionsObjectDefinitionIds

                    if (containedDocumentModel && !containedDocumentAlreadyInFullCollections) {
                        foundedDocumentCollections.add(new DocumentCollectionInfo(containedDocument, containedDocumentModel))
                    }
                }
            }
            return foundedDocumentCollections
        }
    }

    private class ScenarioDescription {
        EPCDescription scenario
        List<ProcedureDescription> procedures = []

        ScenarioDescription(Model model, ObjectElement functionObject) {
            this.scenario = new EPCDescription(model, functionObject)
        }

        ScenarioDescription(Model model) {
            this.scenario = new EPCDescription(model)
        }

        private void defineProcedures() {
            List<ObjectElement> procedureObjects = scenario.model.findObjectsByType(FUNCTION_OBJECT_TYPE_ID)
                    .unique(Comparator.comparing { ObjectElement o -> o.getId() })
                    .sort {o1, o2 -> ModelUtils.getElementsCoordinatesComparator().compare(o1, o2)}

            procedureObjects.each {ObjectElement procedureObject ->
                Model procedureModel = getEPCModel(procedureObject)

                if (procedureModel == null) {
                    // TODO: что делать, если выбран режим до 4 уровня, а у какого-либо 3 лвла нет декомпозиции на 4?
                    return
                }

                procedures.add(new ProcedureDescription(procedureModel, procedureObject))
            }
        }

        private void defineBusinessRoles() {
            procedures.each {it.findBusinessRoles()}
        }

        private List<BusinessRoleInfo> getAllRoles() {
            List<BusinessRoleInfo> allRoles = []
            procedures.each {ProcedureDescription procedure ->
                allRoles.addAll(procedure.businessRoles)
            }
            return allRoles.unique(Comparator.comparing { BusinessRoleInfo bRI -> bRI.businessRole.object.getObjectDefinitionId() })
        }
    }

    private class ProcedureDescription {
        EPCDescription procedure
        List<BusinessRoleInfo> businessRoles = []

        ProcedureDescription(Model model, ObjectElement functionObject) {
            this.procedure = new EPCDescription(model, functionObject)
        }

        private void findBusinessRoles() {
            List<ObjectElement> businessRoleObjects = procedure.model.findObjectsByType(BUSINESS_ROLE_OBJECT_TYPE_ID)
                    .unique(Comparator.comparing { ObjectElement o -> o.getObjectDefinitionId() })
                    .sort {o1, o2 -> ModelUtils.getElementsCoordinatesComparator().compare(o1, o2)}
            businessRoles = businessRoleObjects.collect {new BusinessRoleInfo(it)}
        }
    }

    private class EPCDescription {
        CommonFunctionInfo functionInfo
        Model model
        List<NormativeDocumentInfo> normativeDocuments = []
        List<DocumentCollectionInfo> documentCollections = []

        EPCDescription(Model model, ObjectElement functionObject) {
            this.functionInfo = new CommonFunctionInfo(functionObject)
            this.model = model
        }

        EPCDescription(Model model) {
            this.functionInfo = new CommonFunctionInfo(model)
            this.model = model
        }

        private void findNormativeDocuments () {
            List<ObjectElement> normativeDocumentObjects = model.getObjects()
                    .findAll {it.getSymbolId() == NORMATIVE_DOCUMENT_SYMBOL_ID}
                    .unique(Comparator.comparing { ObjectElement o -> o.getObjectDefinitionId() })
            normativeDocuments = normativeDocumentObjects.collect {new NormativeDocumentInfo(it)}
        }

        private void findDocumentCollections () {
            List<ObjectElement> documentCollectionObjects = model.getObjects()
                    .findAll { it.getObjectDefinition().getObjectTypeId() in DOCUMENT_COLLECTION_OBJECT_TYPE_IDS }
                    .unique(Comparator.comparing { ObjectElement o -> o.getObjectDefinitionId() })
                    .sort { o1, o2 -> ModelUtils.getElementsCoordinatesComparator().compare(o1, o2) }

            documentCollectionObjects.each {ObjectElement documentCollectionObject ->
                Model documentCollectionModel = findDocumentCollectionModel(documentCollectionObject)

                // TODO: обсудить логику определения набора документов (пример ошибки с типом символа)
                if (documentCollectionModel) {
                    documentCollections.add(new DocumentCollectionInfo(documentCollectionObject, documentCollectionModel))
                }
            }
        }

        static Model findDocumentCollectionModel(ObjectElement documentCollectionObject) {
            List <Model> documentCollectionObjectModels = documentCollectionObject.getDecompositions()
                    .findAll {it.isModel()} as List <Model>
            return documentCollectionObjectModels
                    .findAll {it.getModelTypeId() in DOCUMENT_COLLECTION_MODEL_TYPE_IDS}
                    .stream()
                    .findFirst()
                    .orElse(null)
        }
    }

    private class SubprocessOwnerDescription {
        CommonObjectInfo subprocessOwner
        SubprocessOwnerType type
        String leadershipPosition = null

        SubprocessOwnerDescription(ObjectElement subprocessOwner, SubprocessOwnerType subprocessOwnerType) {
            this.subprocessOwner = new CommonObjectInfo(subprocessOwner)
            this.type = subprocessOwnerType

            if (type == SubprocessOwnerType.ORGANIZATIONAL_UNIT) {
                defineLeadershipPosition()
            }
        }

        private void defineLeadershipPosition() {
            ObjectDefinition subprocessOwnerObjectDefinition = subprocessOwner.object.getObjectDefinition()
            Model subprocessOwnerModel = subprocessOwnerObjectDefinition
                    .getDecompositions(ORGANIZATION_STRUCTURE_MODEL_TYPE_ID)
                    .stream()
                    .findFirst()
                    .orElse(null)

            if (subprocessOwnerModel == null) {
                return
            }

            ObjectElement subprocessOwnerModelObject = subprocessOwnerModel.findObjectInstances(subprocessOwnerObjectDefinition)
                    .stream()
                    .findFirst()
                    .orElse(null)

            if (subprocessOwnerModelObject == null) {
                return
            }

            ObjectElement leadershipPositionObject = subprocessOwnerModelObject.getEnterEdges()
                    .find {it.getEdgeTypeId() == LEADERSHIP_POSITION_W_OWNER_EDGE_TYPE_ID}
                    .getSource() as ObjectElement
            this.leadershipPosition = getName(leadershipPositionObject.getObjectDefinition())
        }
    }

    private class InputFlowDescription {
        CommonObjectInfo inputFlow
        List<CommonFunctionInfo> suppliers = []

        InputFlowDescription(ObjectElement inputFlow, List<ObjectElement> supplierObjects) {
            this.inputFlow = new CommonObjectInfo(inputFlow)
            this.suppliers = supplierObjects.collect {new CommonFunctionInfo(it)}
        }
    }

    private class OutputFlowDescription {
        CommonObjectInfo outputFlow
        public List<CommonFunctionInfo> customers = []

        OutputFlowDescription(ObjectElement outputFlow, List<ObjectElement> customerObjects) {
            this.outputFlow = new CommonObjectInfo(outputFlow)
            this.customers = customerObjects.collect {new CommonFunctionInfo(it)}
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

    private static void findAbbreviations(String name) {
        Matcher matcher = abbreviationsPattern.matcher(name)
        while (matcher.find()) {
            String abbreviationName = name.substring(matcher.start(), matcher.end())

            if (abbreviationName in foundedAbbreviations.keySet()) {
                continue
            }

            String abbreviationDescription = fullAbbreviations.get(abbreviationName)
            foundedAbbreviations.put(abbreviationName, abbreviationDescription)
        }
    }

    private static String getAttributeValue(TreeNode treeNode, String attributeId) {
        Node node = treeNode._getNode()
        AttributeValue attribute = node.getAttributes().stream()
                .filter { it.typeId == attributeId}
                .findFirst()
                .orElse(null)

        if (attribute != null && attribute.value != null && !attribute.value.trim().isEmpty()) {
            return trimStringValue(attribute.value)
        }

        return ''
    }

    private static String trimStringValue(String value) {
        String resultString = value.replaceAll("\\u00A0", " ")
        resultString = resultString.replaceAll("[\\s\\n]+", " ").trim()
        return resultString
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

        List<ObjectElement> subProcessObjects = getSubProcessObjects()
        List<SubprocessDescription> subProcessDescriptions = getSubProcessDescriptions(subProcessObjects)
    }

    private void init() {
        treeRepository = context.createTreeRepository(true)
        parseParameters()
        initAbbreviations()
    }

    private void parseParameters() {
        if (debug) {
            detailLevel = 4
            docVersion = '1.0.0'
            docDate = '01.01.2025'
            return
        }

        String deep = ParamUtils.parse(context.findParameter(DETAIL_LEVEL_PARAM_NAME)) as String
        detailLevel = Integer.parseInt(deep.replaceAll("[^0-9]+", ""))

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
                .find {it.getObjectDefinitionId() == ABBREVIATIONS_ROOT_OBJECT_ID}

        if (!abbreviationsRootObject) {
            throw new SilaScriptException("Неверный ID корневого объекта аббревиатур [${ABBREVIATIONS_ROOT_OBJECT_ID}]")
        }

        List<ObjectElement> abbreviationObjects = abbreviationsRootObject.getExitEdges()
                .findAll {it.getEdgeTypeId() in ABBREVIATIONS_EDGE_TYPE_IDS}
                .collect {it.getTarget() as ObjectElement}
                .unique(Comparator.comparing { ObjectElement o -> o.getId() })

        abbreviationObjects.addAll(
                abbreviationsRootObject.getEnterEdges()
                        .findAll {it.getEdgeTypeId() in ABBREVIATIONS_EDGE_TYPE_IDS}
                        .collect {it.getSource() as ObjectElement}
                        .unique(Comparator.comparing { ObjectElement o -> o.getId() })
        )

        for (abbreviationObject in abbreviationObjects) {
            ObjectDefinitionNode abbreviationObjectDefinitionNode = abbreviationObject.getObjectDefinition()._getNode() as ObjectDefinitionNode

            String abbreviationName = abbreviationObjectDefinitionNode.getName()
            String abbreviationDescription = ''
            AttributeValue descriptionDefinitionAttribute = abbreviationObjectDefinitionNode.getAttributes().stream()
                    .filter { it.typeId == DESCRIPTION_DEFINITION_ATTR_ID}
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

    private List<ObjectElement> getSubProcessObjects() {
        List<ObjectElement> subProcessObjects = []
        if (!context.elementsIdsList().isEmpty()){
            Model model = treeRepository.read(context.modelId().getRepositoryId(), context.modelId().getId())
            List<ObjectElement> allObjects = model.getObjects()
            for (elementId in context.elementsIdsList()) {
                for (object in allObjects) {
                    if (object.getId() == elementId) {
                        subProcessObjects.add(object)
                        break
                    }
                }
            }
            subProcessObjects.sort {o1, o2 -> ModelUtils.getElementsCoordinatesComparator().compare(o1, o2)}
        }
        if (subProcessObjects.isEmpty()) {
            throw new SilaScriptException("Скрипт должен запускаться на экземплярах объектов")
        }
        return subProcessObjects
    }

    private List<SubprocessDescription> getSubProcessDescriptions(List<ObjectElement> subProcessObjects) {
        List<SubprocessDescription> subProcessDescriptions = subProcessObjects.collect{new SubprocessDescription(it, detailLevel)}
        subProcessDescriptions.each {buildSubProcessDescription(it)}
        return subProcessDescriptions
    }

    private void buildSubProcessDescription(SubprocessDescription subprocessDescription) {
        subprocessDescription.defineParentProcess()
        subprocessDescription.findOwners()
        subprocessDescription.defineGoals()
        subprocessDescription.findExternalProcessInputFlows()
        subprocessDescription.findExternalProcessOutputFlows()
        subprocessDescription.buildExternalProcessesWithInputFlows()
        subprocessDescription.buildExternalProcessesWithOutputFlows()
        subprocessDescription.defineProcessSelectionModel()
        subprocessDescription.defineScenarios()

        if (detailLevel == 4) {
            subprocessDescription.defineProcedures()
            subprocessDescription.defineBusinessRoles()
        }

        subprocessDescription.identifyAnalyzedEPC()
        subprocessDescription.defineNormativeDocuments()
        subprocessDescription.defineDocumentCollections()
        subprocessDescription.buildFullDocumentCollections()
    }
}
