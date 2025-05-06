import groovy.util.logging.Slf4j
import ru.nextconsulting.bpm.dto.NodeId
import ru.nextconsulting.bpm.repository.business.AttributeValue
import ru.nextconsulting.bpm.repository.structure.ObjectDefinitionNode
import ru.nextconsulting.bpm.repository.structure.ScriptParameter
import ru.nextconsulting.bpm.repository.structure.SilaScriptParamType
import ru.nextconsulting.bpm.script.repository.TreeRepository
import ru.nextconsulting.bpm.script.tree.elements.ObjectElement
import ru.nextconsulting.bpm.script.tree.node.Model
import ru.nextconsulting.bpm.script.tree.node.ObjectDefinition
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
    private static final List<String> ABBREVIATIONS_EDGE_TYPE_IDS = [
            'CT_IS_IN_RELSHP_TO_1',
            'CT_REFS_TO_2',
            'CT_HAS_REL_WITH',
            'CT_IS_IN_RELSHP_TO',
    ]
    private static Map<String, String> fullAbbreviations = new TreeMap<>()
    private static Pattern abbreviationsPattern = null
    private static Map<String, String> foundedAbbreviations = new TreeMap<>()

    private static final String FULL_NAME_ATTR_ID = 'AT_NAME_FULL'
    private static final String DATA_ELEMENT_CODE_ATTR_ID = '46e148b0-b96d-11e3-05b7-db7cafd96ef7'
    private static final String DESCRIPTION_DEFINITION_ATTR_ID = 'AT_DESC'

    CustomScriptContext context
    private TreeRepository tree_repository

    private static int detailLevel = 3
    private static String docVersion = ''
    private static String docDate = ''

    private static boolean debug = true

    private static String getName(ObjectDefinitionNode objectDefinitionNode) {
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
        } else {
            return ''
        }
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

    private static String getCode(ObjectDefinitionNode objectDefinitionNode) {
        AttributeValue codeAttribute = objectDefinitionNode.getAttributes().stream()
                .filter { it.typeId == DATA_ELEMENT_CODE_ATTR_ID}
                .findFirst()
                .orElse(null)

        if (codeAttribute != null && codeAttribute.value != null && !codeAttribute.value.trim().isEmpty()) {
            return codeAttribute.value
        }

        return ''
    }


    private class SubprocessDescription {
        ObjectElement subprocess
        String name
        String code


        SubprocessDescription(ObjectElement subprocess) {
            this.subprocess = subprocess

            ObjectDefinition objectDefinition = subprocess.getObjectDefinition()
            ObjectDefinitionNode objectDefinitionNode = subprocess.getObjectDefinition()._getNode() as ObjectDefinitionNode

            String name = getName(objectDefinitionNode)
            if (!name) {
                name = '<Наименование процесса>'
            }
            this.name = name

            String code = getCode(objectDefinitionNode)
            if (!code) {
                code = '<Код процесса>'
            }
            this.code = code


        }
    }

    @Override
    void execute() {
        init()

        List<ObjectElement> subProcessObjects = getSubProcessObjects()
        List<SubprocessDescription> subProcessDescriptions = getSubProcessDescriptions(subProcessObjects)
    }

    private void init() {
        tree_repository = context.createTreeRepository(true)
        parse_parameters()
        init_abbreviations()
    }

    private void parse_parameters() {
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

    private void init_abbreviations() {
        Model abbreviationsModel = tree_repository.read(context.modelId().getRepositoryId(), ABBREVIATIONS_MODEL_ID)
        if (!abbreviationsModel) {
            throw new SilaScriptException("Неверный ID модели аббревиатур [${ABBREVIATIONS_MODEL_ID}]")
        }

        List<ObjectElement> allObjects = abbreviationsModel.getObjects()
        ObjectElement abbreviationsRootObject = null
        for (object in allObjects) {
            if (object.getObjectDefinition().getId() == ABBREVIATIONS_ROOT_OBJECT_ID) {
                abbreviationsRootObject = object
                break
            }
        }

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
            AttributeValue descriprionDefinitionAttribute = abbreviationObjectDefinitionNode.getAttributes().stream()
                    .filter { it.typeId == DESCRIPTION_DEFINITION_ATTR_ID}
                    .findFirst()
                    .orElse(null)
            if (descriprionDefinitionAttribute != null && descriprionDefinitionAttribute.value != null && !descriprionDefinitionAttribute.value.trim().isEmpty()) {
                abbreviationDescription = descriprionDefinitionAttribute.value
            }
            fullAbbreviations.put(abbreviationName, abbreviationDescription)
        }

        Set<String> abbreviationNames = fullAbbreviations.keySet()
        abbreviationsPattern = Pattern.compile("\\b(?:(?:${String.join(')|(?:', abbreviationNames)}))\\b")
    }

    private List<ObjectElement> getSubProcessObjects() {
        List<ObjectElement> subProcessObjects = []
        if (!context.elementsIdsList().isEmpty()){
            Model model = tree_repository.read(context.modelId().getRepositoryId(), context.modelId().getId())
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

        return subProcessDescriptions
    }
}
