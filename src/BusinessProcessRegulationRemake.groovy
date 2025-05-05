import groovy.util.logging.Slf4j
import ru.nextconsulting.bpm.dto.NodeId
import ru.nextconsulting.bpm.repository.structure.ScriptParameter
import ru.nextconsulting.bpm.repository.structure.SilaScriptParamType
import ru.nextconsulting.bpm.scriptengine.context.ContextParameters
import ru.nextconsulting.bpm.scriptengine.context.CustomScriptContext
import ru.nextconsulting.bpm.scriptengine.script.GroovyScript
import ru.nextconsulting.bpm.scriptengine.util.SilaScriptParameter
import ru.nextconsulting.bpm.scriptengine.util.SilaScriptParameters
import ru.nextconsulting.bpm.utils.JsonConverter

void execute() {
    new BusinessProcessRegulationRemakeScript(context: context).execute()
}

@SilaScriptParameters([
        @SilaScriptParameter(
                name = 'Глубина детализации регламента',
                type = SilaScriptParamType.SELECT_STRING,
                selectStringValues = ['3 уровень', '4 уровень'],
                defaultValue = '3 уровень'
        ),
        @SilaScriptParameter(
                name = 'Номер версии регламента',
                type = SilaScriptParamType.STRING,
                required = true
        ),
        @SilaScriptParameter(
                name = 'Дата утверждения регламента',
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
                        .id('d047bce0-dcc1-11e6-05b7-db7cafd96ef7')
                        .build())
                )
                .build()
        ScriptParameter elementsIdsParam = ScriptParameter.builder()
                .paramType(SilaScriptParamType.STRING_LIST)
                .name('elementsIdsList')
                .value('["5fa0d203-dcc2-11e6-05b7-db7cafd96ef7"]')
                .build()

        context.getParameters().add(modelParam)
        context.getParameters().add(elementsIdsParam)

        BusinessProcessRegulationRemakeScript script = new BusinessProcessRegulationRemakeScript(context: context)
        script.execute()
    }

    CustomScriptContext context

    @Override
    void execute() {

    }
}
