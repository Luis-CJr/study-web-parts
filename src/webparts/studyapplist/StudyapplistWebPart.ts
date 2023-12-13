import { Version } from '@microsoft/sp-core-library';
import { IPropertyPaneConfiguration, PropertyPaneTextField } from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import * as strings from 'StudyapplistWebPartStrings';
import * as $ from "jquery";
import { sp } from "@pnp/sp";
import "bootstrap";
import Swal from 'sweetalert2'
import 'sweetalert2/dist/sweetalert2.min.css'
require('../../../node_modules/bootstrap/dist/css/bootstrap.min.css');
require('../../../node_modules/@fortawesome/fontawesome-free/css/all.min.css');
require('../../stylelibrary/css/padrao.css');
require('../../stylelibrary/css/toastr.min.css');

export interface IStudyapplistWebPartProps {
  description: string;
  usuarioAtual: string;
}

export default class StudyapplistWebPart extends BaseClientSideWebPart<IStudyapplistWebPartProps> {
  public usuarioAtual: string;

  public async onInit(): Promise<void> {
    await super.onInit();
    sp.setup({
      spfxContext: this.context
    });
    const user = await sp.web.currentUser.get();
    this.usuarioAtual = user.Title;
  };

  public render(): void {
    this.domElement.innerHTML = require('./list.html');
    let usuarioAtivo = document.getElementById("usuarioLogado");
    usuarioAtivo.textContent = this.usuarioAtual;

    const eventHandlers = new EventHandlers(this.domElement, this.usuarioAtual);
    eventHandlers.setEventHandlers();
  };

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  };

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: strings.PropertyPaneDescription
          },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneTextField('description', {
                  label: strings.DescriptionFieldLabel
                })
              ]
            }
          ]
        }
      ]
    };
  }
}

class EventHandlers {
  private domElement: HTMLElement;
  private usuarioAtual: string;

  constructor(domElement: HTMLElement, usuarioAtual: string) {
    this.domElement = domElement;
    this.usuarioAtual = usuarioAtual;
  };

  public abrirModal(): void {
    $('#modalAtualizar').modal('show')
  };

  public setEventHandlers(): void {
    const btnVerTudo = document.getElementById('btnVerTudo');

    if (btnVerTudo) {
      btnVerTudo.addEventListener('click', () => {
        this.abrirModal();
        this.getEstudosDoUsuario(this.usuarioAtual);
      });
    }
  };

  public async getEstudosDoUsuario(usuarioAtual: string): Promise<void> {
    let dados = await sp.web.lists.getByTitle("estudosDoUsuario").items
      .filter(`Title eq '${usuarioAtual}'`)
      .select("ID", "Title", "datainicio", "datafim", "conteudo/Title", "descricao/descricao", "datafim")
      .expand("conteudo", "descricao")
      .get();

    console.log(dados)
    this.renderItemsInTable(dados);
  };

  public renderItemsInTable(items: any[]): void {
    let tableBodyHTML = "";

    items.forEach((item) => {
      let contentValue = item.conteudo.Title;
      let descriptValue = item.descricao.descricao;
      let finished = item.datafim ? new Date(item.datafim).toLocaleDateString() : "Não finalizado";

      let finishButtonHTML = item.datafim ? "" : `<button class="btn btn-primary btn-sm finishBtn" data-item-id="${item.Id}">Finalizar</button>`;
      let deleteButtonHTML = `<button class="btn btn-primary btn-sm deleteBtn" data-item-id="${item.Id}">Excluir</button>`;

      tableBodyHTML += `<tr id="itemRow-${item.Id}">
          <td>${contentValue}</td>
          <td>${descriptValue}</td>
          <td>${new Date(item.datainicio).toLocaleDateString()}</td>
          <td class="datafim">${finished}</td>
          <td>${finishButtonHTML} ${deleteButtonHTML}</td>
      </tr>`;
    });

    $("#modalEstudosContainer").html(tableBodyHTML);

    $('.finishBtn').on('click', (event) => {
      let itemId = $(event.currentTarget).attr('data-item-id');
      this.Finish(Number(itemId));
    });

    $('.deleteBtn').on('click', (event) => {
      let itemId = $(event.currentTarget).attr('data-item-id');
      this.Delete(Number(itemId))
    });
  };

  private async Finish(itemId: number): Promise<void> {
    try {
      const hoje = new Date();
      await sp.web.lists.getByTitle("estudosDoUsuario").items.getById(itemId).update({
        datafim: hoje,
      });
      // alert("Parabéns, estudo finalizado!");
      Swal.fire({
        title: 'Parabéns!',
        text: 'Trilha de estudos concluída.',
        icon: 'success',
        timer: 3000
      });
      this.atualizarItemNaUI(itemId, hoje);
    } catch (error) {
      console.error("Erro ao atualizar data:", error);
    }
  };

  private atualizarItemNaUI(itemId: number, dataFim: Date): void {
    const dataFormatada = dataFim.toLocaleDateString();

    const linhaItem = $(`#itemRow-${itemId}`);
    linhaItem.find('.datafim').text(dataFormatada);
    linhaItem.find('.finishBtn').hide();
  };

  private async Delete(itemId: number): Promise<void> {
    try {
      await sp.web.lists.getByTitle("estudosDoUsuario").items.getById(itemId).delete();
      // alert("Exclusão feita com sucesso.");
      Swal.fire({
        text: 'Item excluído.',
        timer: 3000,
      });
      this.removerItemDaUI(itemId)
    } catch (error) {
      console.error("Erro ao excluir item:", error);
    }
  };

  private removerItemDaUI(itemId: number): void {
    $(`#itemRow-${itemId}`).remove();
  };
}