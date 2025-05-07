import groovy.util.logging.Slf4j
import ru.nextconsulting.bpm.dto.NodeId
import ru.nextconsulting.bpm.repository.business.AttributeValue
import ru.nextconsulting.bpm.repository.structure.ObjectDefinitionNode
import ru.nextconsulting.bpm.repository.structure.ScriptParameter
import ru.nextconsulting.bpm.repository.structure.SilaScriptParamType
import ru.nextconsulting.bpm.script.repository.TreeRepository
import ru.nextconsulting.bpm.script.tree.elements.ObjectElement
import ru.nextconsulting.bpm.script.tree.node.Model
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

    private static final GROUP_OBJECT_TYPE_ID = 'OT_GRP'
    private static final ORGANIZATIONAL_UNIT_OBJECT_TYPE_ID = 'OT_ORG_UNIT'

    private static final List<String> ABBREVIATIONS_EDGE_TYPE_IDS = [
            'CT_IS_IN_RELSHP_TO_1',
            'CT_REFS_TO_2',
            'CT_HAS_REL_WITH',
            'CT_IS_IN_RELSHP_TO',
    ]
    private static final String LEADERSHIP_POSITION_W_OWNER_EDGE_TYPE_ID = 'CT_IS_DISC_SUPER'
    private static final List<String> OWNER_W_SUBPROCESS_EDGE_TYPE_IDS = [
            'CT_EXEC_1',
            'CT_EXEC_2',
    ]

    private static final String DATA_ELEMENT_CODE_ATTR_ID = '46e148b0-b96d-11e3-05b7-db7cafd96ef7'
    private static final String DESCRIPTION_DEFINITION_ATTR_ID = 'AT_DESC'
    private static final String FULL_NAME_ATTR_ID = 'AT_NAME_FULL'

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

    private class SubprocessDescription {
        ObjectElement subprocess
        String name
        String code
        List<SubprocessOwnerDescription> owners = []

        SubprocessDescription(ObjectElement subprocess) {
            this.subprocess = subprocess

            String name = getName(subprocess)
            if (!name) {
                name = '<Наименование процесса>'
            }
            this.name = name

            String code = getCode(subprocess)
            if (!code) {
                code = '<Код процесса>'
            }
            this.code = code
        }

        void findOwners() {
            List<ObjectElement> ownerObjects = subprocess.getEnterEdges()
                    .findAll {it.getEdgeTypeId() in OWNER_W_SUBPROCESS_EDGE_TYPE_IDS}
                    .collect {it.getSource() as ObjectElement}
                    .unique(Comparator.comparing { ObjectElement o -> o.getId() })
                    .sort {o1, o2 -> ModelUtils.getElementsCoordinatesComparator().compare(o1, o2)}
            owners = ownerObjects.collect {
                new SubprocessOwnerDescription(it, subprocessOwnerTypeMap.get(it.getObjectDefinition().getObjectTypeId()))
            }
        }
    }

    private class SubprocessOwnerDescription {
        ObjectElement subprocessOwner
        String name
        SubprocessOwnerType type
        String leadershipPosition = null

        SubprocessOwnerDescription(ObjectElement subprocessOwner, SubprocessOwnerType subprocessOwnerType) {
            this.subprocessOwner = subprocessOwner

            String name = getName(subprocessOwner)
            if (!name) {
                name = '<Владелец процесса>'
            }
            this.name = name

            this.type = subprocessOwnerType

            if (type == SubprocessOwnerType.ORGANIZATIONAL_UNIT) {
                defineLeadershipPosition()
            }
        }

        private void defineLeadershipPosition() {
            String subprocessOwnerObjectDefinitionId = subprocessOwner.getObjectDefinition().getId()
            Model subprocessOwnerModel = subprocessOwner.getDecompositions()
                    .findAll {it.isModel()}
                    .stream()
                    .findFirst()
                    .orElse(null)
                    as Model

            if (subprocessOwnerModel == null) {
                return
            }

            ObjectElement subprocessOwnerModelObject = subprocessOwnerModel.getObjects()
                    .find {it.getObjectDefinition().getId() == subprocessOwnerObjectDefinitionId}

            if (subprocessOwnerModelObject == null) {
                return
            }

            ObjectElement leadershipPositionObject = subprocessOwnerModelObject.getEnterEdges()
                .find {it.getEdgeTypeId() == LEADERSHIP_POSITION_W_OWNER_EDGE_TYPE_ID}
                .getSource() as ObjectElement
            String position = getName(leadershipPositionObject)
            if (!position) {
                position = '<Должность владельца процесса>'
            }
            this.leadershipPosition = position
        }
    }

    private static String getName(ObjectElement objectElement) {
        ObjectDefinitionNode objectDefinitionNode = objectElement.getObjectDefinition()._getNode() as ObjectDefinitionNode
        String name = objectDefinitionNode.getName()

        if (name) {
            findAbbreviations(name)
        }

        AttributeValue fullNameAttribute = objectDefinitionNode.getAttributes().stream()
                .filter { it.typeId == FULL_NAME_ATTR_ID }
                .findFirst()
                .orElse(null)
        if (fullNameAttribute != null && fullNameAttribute.value != null && !fullNameAttribute.value.trim().isEmpty()) {
            name = fullNameAttribute.value
            findAbbreviations(name)
        }

        if (name) {
            return name.replaceAll("[\\s\\n]+", " ").trim()
        }

        return ''
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

    private static String getCode(ObjectElement objectElement) {
        ObjectDefinitionNode objectDefinitionNode = objectElement.getObjectDefinition()._getNode() as ObjectDefinitionNode
        AttributeValue codeAttribute = objectDefinitionNode.getAttributes().stream()
                .filter { it.typeId == DATA_ELEMENT_CODE_ATTR_ID}
                .findFirst()
                .orElse(null)

        if (codeAttribute != null && codeAttribute.value != null && !codeAttribute.value.trim().isEmpty()) {
            return codeAttribute.value
        }

        return ''
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
            detailLevel = 3
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

        List<ObjectElement> allObjects = abbreviationsModel.getObjects()
        ObjectElement abbreviationsRootObject = allObjects
                .find {it.getObjectDefinition().getId() == ABBREVIATIONS_ROOT_OBJECT_ID}

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
        List<SubprocessDescription> subProcessDescriptions = subProcessObjects.collect{new SubprocessDescription(it)}
        subProcessDescriptions.each {it.findOwners()}


        return subProcessDescriptions
    }
}
