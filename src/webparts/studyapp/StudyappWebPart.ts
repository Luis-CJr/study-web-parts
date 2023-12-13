import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import * as strings from 'StudyappWebPartStrings';
import * as $ from "jquery"
import { sp } from "@pnp/sp";
import "bootstrap";
import Swal from 'sweetalert2'
import 'sweetalert2/dist/sweetalert2.min.css'
require('../../../node_modules/bootstrap/dist/css/bootstrap.min.css');
require('../../../node_modules/@fortawesome/fontawesome-free/css/all.min.css');
require('../../stylelibrary/css/padrao.css');
require('../../stylelibrary/css/toastr.min.css');

export interface IStudyappWebPartProps {
  description: string;
}

export default class StudyappWebPart extends BaseClientSideWebPart<IStudyappWebPartProps> {
  private usuarioAtual: string;

  public async onInit(): Promise<void> {

    return super.onInit().then(_ => {
      sp.setup({
        spfxContext: this.context
      })
    });

  }

  public render(): void {
    this.domElement.innerHTML = require('./form.html');
    this.fillDropdown();

    const eventHandlers = new EventHandlers(this.domElement, this.usuarioAtual);
    eventHandlers.setEventListeners();
  }

  private fillDropdown(): void {
    sp.web.lists.getByTitle("conteudoDeEstudos").items.select("Title").get().then((items) => {
      let dropdown = $("#ddlMateriaEstudo").empty();

      dropdown.append($("<option>", {
        value: "",
        text: "Conteúdo",
        disable: true,
        selected: true

      }))

      items.forEach((item) => {
        $("#ddlMateriaEstudo").append($("<option>", {
          value: item.Title,
          text: item.Title,
        }));
      });
    })
      .catch((error) => {
        console.error("Erro ao carregar opções do dropdown: " + error);
      });
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

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
  }

  public setEventListeners(): void {
    const btnConfirmarEstudo = this.domElement.querySelector('#btnConfirmarEstudo');
    if (btnConfirmarEstudo) {
      btnConfirmarEstudo.addEventListener('click', () => {
        this.handleConfirmarEstudoClick();
      });
    }
  };

  private handleConfirmarEstudoClick(): void {
    const materiaEstudo: string = (this.domElement.querySelector('#ddlMateriaEstudo') as HTMLSelectElement).value;
    const dataEstudo: string = (this.domElement.querySelector('#txtDataEstudo') as HTMLInputElement).value;

    if (materiaEstudo && dataEstudo) {
      this.salvarEstudoUsuario(materiaEstudo, dataEstudo, this.usuarioAtual);
    } else {
      // alert('Preencha todos os campos antes de confirmar o estudo.');
      Swal.fire({
        title: 'Atenção!',
        text: 'Preencha todos os campos antes de confirmar o estudo.',
        icon: 'warning',
        confirmButtonText: 'Ok'
      });
    }
  }

  private async salvarEstudoUsuario(materiaEstudo: string, dataEstudo: string, _usuarioAtual: string): Promise<void> {
    try {
      sp.web.currentUser.get().then(user => {
        this.usuarioAtual = user.Title;
      });

      const conteudos = await sp.web.lists.getByTitle("conteudoDeEstudos").items
        .select("Id", "Title", "descricao")
        .filter(`Title eq '${materiaEstudo}'`)
        .get();

      if (conteudos.length === 0) throw new Error("Conteúdo de estudo não encontrado.");

      const novoEstudo = {
        Title: this.usuarioAtual,
        datainicio: dataEstudo,
        conteudoId: conteudos[0].Id,
        descricaoId: conteudos[0].Id
      };

      await sp.web.lists.getByTitle("estudosDoUsuario").items.add(novoEstudo);
      // alert("Estudo confirmado com sucesso!");
      Swal.fire({
        title: 'Sucesso!',
        text: 'O item foi adicionado à sua lista de estudos.',
        icon: 'success',
        timer: 3000
      });

    } catch (error) {
      console.error("Erro ao salvar estudo do usuário: ", error);
      // alert("Erro ao salvar estudo do usuário.");
      Swal.fire({
        title: 'Ops...!',
        text: 'Não foi possível adicionar o item a sua lista de estudos.',
        icon: 'warning',
        timer: 3000,
        confirmButtonText: 'Tentar novamente'
      });
    }
  }
}