// This file is part of Acta Nova (www.acta-nova.eu)
// Copyright (c) rubicon IT GmbH, www.rubicon.eu
// Version 1.1 - LHIND - Philipp Rössler - 05.09.2019

using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using ActaNova.Domain;
using ActaNova.Domain.Classes.Configuration;
using ActaNova.Domain.Classifications;
using ActaNova.Domain.ExpressionDefinitions;
using ActaNova.Domain.Extensions;
using ActaNova.Domain.Specialdata;
using ActaNova.Domain.Specialdata.Catalog;
using ActaNova.Domain.Specialdata.Values;
using ActaNova.Domain.Testing;
using ActaNova.Domain.Workflow;
using Remotion.Data.DomainObjects;
using Remotion.Data.DomainObjects.DomainImplementation;
using Remotion.Data.DomainObjects.Queries;
using Remotion.Globalization;
using Remotion.Logging;
using Remotion.ObjectBinding;
using Remotion.Security;
using Remotion.SecurityManager.Domain.OrganizationalStructure;
using Rubicon.Dms;
using Rubicon.Domain;
using Rubicon.Domain.TenantBoundOrganizationalStructure;
using Rubicon.Gever.Bund.Domain.Utilities;
using Rubicon.Gever.Bund.Domain.Utilities.Extensions;
using Rubicon.Multilingual.Extensions;
using Rubicon.ObjectBinding.Utilities;
using Rubicon.Utilities;
using Rubicon.Utilities.Globalization;
using Rubicon.Utilities.Security;
using Rubicon.Workflow.Domain;
using BusinessObjectExtensions = Remotion.ObjectBinding.BusinessObjectExtensions;

namespace LHIND.Mitbericht
{
  [LocalizationEnum]
  public enum WorkflowLocalization
  {
    [MultiLingualName("Das Geschäftsobjekt der Aktivität \"{0}\" muss ein Dossier sein", "")]
    NoFile,

    [MultiLingualName("Das Geschäftsobjekt der Aktivität \"{0}\" muss ein Geschäftsvorfall sein", "")]
    NoFileCase,

    [MultiLingualName("Das Geschäftsobjekt \"{0}\" der Aktivität kann nicht bearbeitet werden", "")]
    NotEditable,

    [MultiLingualName("Es wurde keine Geschäftsobjektvorlage mit der Referenz-ID \"{0}\" gefunden", "")]
    NoFileCaseTemplate,

    [MultiLingualName("Es wurde noch keine Federführung Amt definiert.", "")]
    NoFFDefined,

    [MultiLingualName("Es wurde keine Gruppe fär den Katalogwert \"{0}\" gefunden", "")]
    NoGroupDefined,

    [MultiLingualName("Es wurde keine Geschäftsart \"{0}\" gefunden.", "")]
    NoFileCaseType,

    [MultiLingualName("Es wurden keine Empänger \"{0}\" gefunden.", "")]
    NoRecipients,

    }

  public class LhindMitberichtCreateFilecasesAndAssignActivityCommandModules : ActivityCommandModule
  {
    private static readonly ILog s_logger = LogManager.GetLogger(typeof(LhindMitberichtCreateFilecasesAndAssignActivityCommandModules));

    public LhindMitberichtCreateFilecasesAndAssignActivityCommandModules()
        : base("LHIND_Mitbericht_GVF_GS_erzeugen:ActivityCommandClassificationType")
    {
    }

    public override bool Execute(CommandActivity commandActivity)
    {
      var FileCaseTemplateReferenceID_FF = "6220CB49-8E09-4AB7-B4F1-673C6C91CC7E";
      var FileCaseTemplateReferenceID_MB = "56E01B36-E6DE-4DA3-ABE8-95C6551A76D3";

      try
      {
        var f = commandActivity.WorkItem as File;

        if (f == null)
          throw new ActivityCommandException("")
              .WithUserMessage(WorkflowLocalization.NoFile.ToLocalizedName().FormatWith(commandActivity));

        if (!f.CanEdit(true))
          throw new ActivityCommandException("")
              .WithUserMessage(WorkflowLocalization.NotEditable.ToLocalizedName().FormatWith(f));

        //Template for FF
        var fileCaseTemplateFF = QueryFactory.CreateLinqQuery<FileCaseTemplate>().FirstOrDefault(fct => ((IHasReferenceIDMixin) fct).ReferenceID == FileCaseTemplateReferenceID_FF);
        if (fileCaseTemplateFF == null)
          throw new ActivityCommandException("")
              .WithUserMessage(WorkflowLocalization.NoFileCaseTemplate.ToLocalizedName().FormatWith(FileCaseTemplateReferenceID_FF));

        var ffCatalogValue = f.GetProperty("#LHIND_Mitbericht_federfuhrendesAmt") as SpecialdataCatalogValue;
        if (ffCatalogValue == null)
          throw new ActivityCommandException("")
              .WithUserMessage(WorkflowLocalization.NoFFDefined.ToLocalizedName().FormatWith());

        var ffgroup = ffCatalogValue.GetProperty("#LHIND_Mitbericht_SPOC") as TenantGroup;
        if (ffgroup == null)
          throw new ActivityCommandException("")
              .WithUserMessage(WorkflowLocalization.NoGroupDefined.ToLocalizedName().FormatWith(ffCatalogValue));
        //müssen in einer eigenen Transaktion gestartet werden
        var startFileCases = new List<FileCase>();

        var ff_filecase = FileCase.NewObject (f, null, null, fileCaseTemplateFF);
        ff_filecase.LeadingGroup = ffgroup;
        var newTitle = ff_filecase.GetMultilingualValue(fc => fc.Title) + " " + ffgroup.GetMultilingualValue(g => g.ShortName);
        ff_filecase.SetMultilingualValue(fc => fc.Title, newTitle);
        startFileCases.Add(ff_filecase);

        //Template for MB
        var fileCaseTemplateMB = QueryFactory.CreateLinqQuery<FileCaseTemplate>().FirstOrDefault(fct => ((IHasReferenceIDMixin) fct).ReferenceID == FileCaseTemplateReferenceID_MB);
        if (fileCaseTemplateMB == null)
          throw new ActivityCommandException("")
              .WithUserMessage(WorkflowLocalization.NoFileCaseTemplate.ToLocalizedName().FormatWith(FileCaseTemplateReferenceID_MB));

        var mitbeteiligtCatalogValues = f.GetProperty("#LHIND_Mitbericht_mitbeteiligtFDListe") as SpecialdataListPropertyValueCollection;
        
        foreach (var mitbeteiligtCatalogValue in mitbeteiligtCatalogValues)
        {
          var mgroup = mitbeteiligtCatalogValue.WrappedValue.GetProperty ("#LHIND_Mitbericht_SPOC") as TenantGroup;
          if (mgroup == null)
            throw new ActivityCommandException ("")
                .WithUserMessage (WorkflowLocalization.NoGroupDefined.ToLocalizedName().FormatWith (ffCatalogValue));

          var m_filecase = FileCase.NewObject (f, null, null, fileCaseTemplateMB);

          m_filecase.LeadingGroup = mgroup;

          newTitle = m_filecase.GetMultilingualValue (fc => fc.Title) + " " + mgroup.GetMultilingualValue(g=>g.ShortName);
          m_filecase.SetMultilingualValue (fc => fc.Title, newTitle);
          startFileCases.Add (m_filecase);
        }
        
        ClientTransaction.Current.Commit();
        foreach (var startFileCase in startFileCases)
          startFileCase.StartObject();
        ClientTransaction.Current.Commit();
      }
      catch (Exception ex)
      {
        ClientTransaction.Current.Rollback();
        s_logger.Error(ex.Message);
        throw;
      }
      return true;
    }


    public override string Validate(CommandActivity commandActivity)
    {
      return null;
    }

  }

  public class LhindMitberichtCreateNewIncomingFromFilecaseActivityCommandModules : ActivityCommandModule
  {
    private static readonly ILog s_logger = LogManager.GetLogger(typeof(LhindMitberichtCreateNewIncomingFromFilecaseActivityCommandModules));

    public LhindMitberichtCreateNewIncomingFromFilecaseActivityCommandModules()
        : base("LHIND_Mitbericht_Eingang_erzeugen:ActivityCommandClassificationType")
    {
    }

    public override bool Execute(CommandActivity commandActivity)
    {
      var restoreCtxID = ApplicationContext.CurrentID;
      try
      {
        var fc = commandActivity.WorkItem as FileCase;

        if (fc == null)
          throw new ActivityCommandException("")
              .WithUserMessage(WorkflowLocalization.NoFileCase.ToLocalizedName().FormatWith(commandActivity));

        if (!fc.CanEdit(true))
          throw new ActivityCommandException("")
              .WithUserMessage(WorkflowLocalization.NotEditable.ToLocalizedName().FormatWith(fc));
        
        //Get values -> ToTo - Create Loop
        var terminAmt = String.Empty;
        var terminGS = String.Empty;
        var ffAmt = String.Empty;
        var rueckfragen = String.Empty;
        var mbAemter = String.Empty;
        var rueckmeldung = String.Empty;
        var title = String.Empty;
        var auftragsart = String.Empty;

        try
        {
            terminAmt = fc.GetProperty("#LHIND_Mitbericht_BezeichnerDU2").ToString();
        }
        catch (Exception e)
        {
            //No specialdata property available
        }

        try
        {
            terminGS = fc.GetProperty("#LHIND_Mitbericht_BezeichnerDU1").ToString();
        }
        catch (Exception e)
        {
            //No specialdata property available
        }

        try
        {
            ffAmt = (fc.GetProperty("#LHIND_Mitbericht_federfuhrendesAmt") as SpecialdataCatalogValue).DisplayName;
        }
        catch (Exception e)
        {
            //No specialdata property available
        }

        try
        {
            rueckfragen = (fc.GetProperty("#LHIND_Mitbericht_rückfragenAn") as TenantUser).DisplayName;
        }
        catch (Exception e)
        {
            //No specialdata property available
        }

        try
        {
            var mbAemterCatalogValues = fc.GetProperty("#LHIND_Mitbericht_mitbeteiligtFDListe") as SpecialdataListPropertyValueCollection;
            mbAemter = String.Join(", ", mbAemterCatalogValues.ToList().Select(m => m.WrappedValue.GetProperty("DisplayName")).ToList().ToArray());
        }
        catch (Exception e)
        {
            //No specialdata property available
        }

        try
        {
            title = (fc.GetProperty("#LHIND_Mitbericht_titelIDP") as string);
        }
        catch (Exception e)
        {
            //No specialdata property available
        }

        try
        {
            auftragsart = (fc.GetProperty("#LHIND_Mitbericht_auftragsartMitberichtsverfahren") as SpecialdataCatalogValue).DisplayName;
        }
        catch (Exception e)
        {
            //No specialdata property available
        }

        try
        {
            var rueckmeldungCatalogValues = fc.GetProperty("#LHIND_Mitbericht_rückmeldungAn") as SpecialdataListPropertyValueCollection;
            rueckmeldung = String.Join(", ", rueckmeldungCatalogValues.ToList().Select(m => m.WrappedValue.GetProperty("DisplayName")).ToList().ToArray());
        }
        catch (Exception e)
        {
            //No specialdata property available
        }

        var istFederfuehrung = fc.GetProperty("#LHIND_Mitbericht_istFederfuehrung").ToString();
        Boolean boolFF;
        //s_logger.Fatal(istFederfuehrung);
        Boolean.TryParse(istFederfuehrung, out boolFF);
        //s_logger.Fatal(boolFF);
        var fcURL = UrlProvider.Current.GetOpenWorkListItemUrl(fc);

        var fileContentObjects = fc.FileCaseContents.Where(c => c.Document != null).Select(c => c.Document);
        using (ClientTransaction.CreateRootTransaction().EnterDiscardingScope())
        using (TenantSection.SwitchToTenant(UserHelper.Current.Tenant))
        {
          var incoming = Incoming.NewObject();
          ApplicationContext.CurrentID = incoming.ApplicationContextID;
          incoming.LeadingGroup = UserHelper.Current.GetActaNovaUserExtension().StandardGroup != null
              ? UserHelper.Current.GetActaNovaUserExtension().StandardGroup.AsTenantGroup()
              : UserHelper.Current.OwningGroup.AsTenantGroup();

          incoming.Subject = fc.DisplayNameInternal;
          
          //Eingangstyp und Fachdaten
          var incomingTypeReferenceID = "ABC5B6D0-6764-4801-AC36-7B42B80F63D3";
          incoming.IncomingType = QueryFactory.CreateLinqQuery<IncomingTypeClassificationType>().FirstOrDefault(i => ((IHasReferenceIDMixin) i).ReferenceID == incomingTypeReferenceID );

          var ro = new ActaNova.Domain.Specialdata.SpecialdataIgnoreReadOnlySection();
          incoming.SetProperty("#LHIND_Mitbericht_VE_TerminGS", terminGS);
          incoming.SetProperty("#LHIND_Mitbericht_VE_TerminAmt", terminAmt);
          incoming.SetProperty("#LHIND_Mitbericht_VE_Rückmeldung", rueckmeldung);
          incoming.SetProperty("#LHIND_Mitbericht_VE_ffAmt", ffAmt);
          incoming.SetProperty("#LHIND_Mitbericht_VE_AuftragsartMitbericht", auftragsart);
          incoming.SetProperty("#LHIND_Mitbericht_VE_Titel", title);
          incoming.SetProperty("#LHIND_Mitbericht_VE_Mitbeteiligt", mbAemter);
          incoming.SetProperty("#LHIND_Mitbericht_VE_RückfragenAn", rueckfragen);
          incoming.SetProperty("#LHIND_Mitbericht_VE_BaseObjectURL", fcURL);
          incoming.SetProperty("#LHIND_Mitbericht_VE_istFederfuehrung", boolFF);
          ro.Leave();

          incoming.ExternalNumber = fc.FormattedNumber;
          incoming.Remark = fc.WorkInstruction;
          var template = Template.GetTemplateByNameOrGuid("Eingang: Erstellen");
          ((IStartMainProcessDomainObjectMixin)incoming).CreateMainProcessInstance(template);

          foreach (var fileContentObject in fileContentObjects)
          {
            var copyDocument = (IDocument)Document.NewObject(incoming);
            var originalDocument = fileContentObject.PrimaryContent;
            copyDocument.Name = originalDocument.Name;
            using (TenantSection.DisableQueryRestrictions())
            {
              if (originalDocument.HasContent())
              {

                using (var handle = originalDocument.GetContent())
                  copyDocument.SetContent(handle, originalDocument.Extension, originalDocument.MimeType);
              }
              else
              {
                copyDocument.MimeType = originalDocument.MimeType;
                copyDocument.Extension = originalDocument.Extension;
              }
            }
          }
          ClientTransaction.Current.Commit();
        }
      }
      catch (Exception ex)
      {
        s_logger.Error(ex.Message);
        throw;
      }
      finally
      {
        ApplicationContext.CurrentID = restoreCtxID;
      }
      return true;
    }


    public override string Validate(CommandActivity commandActivity)
    {
      return null;
    }

  }

    public class LhindMitberichtCreateInternalFilecases : ActivityCommandModule
    {
        private static readonly ILog s_logger = LogManager.GetLogger(typeof(LhindMitberichtCreateInternalFilecases));

        public LhindMitberichtCreateInternalFilecases()
            : base("LHIND_Mitbericht_Interne_Weiterleitung:ActivityCommandClassificationType")
        {
        }

        public override bool Execute(CommandActivity commandActivity)
        {
            try
            {
                var f = commandActivity.WorkItem as File;

                if (f == null)
                    throw new ActivityCommandException("")
                        .WithUserMessage(WorkflowLocalization.NoFile.ToLocalizedName().FormatWith(commandActivity));

                if (!f.CanEdit(true))
                    throw new ActivityCommandException("")
                        .WithUserMessage(WorkflowLocalization.NotEditable.ToLocalizedName().FormatWith(f));

                var FileCaseTemplateReferenceID = "04248329-15C8-40C8-8888-90DF1C782A56";

                var fileCaseTemplate_WL = QueryFactory.CreateLinqQuery<FileCaseTemplate>().FirstOrDefault(t => ((IHasReferenceIDMixin)t).ReferenceID == FileCaseTemplateReferenceID);

                if (fileCaseTemplate_WL == null)
                    throw new ActivityCommandException("")
                        .WithUserMessage(WorkflowLocalization.NoFileCaseTemplate.ToLocalizedName().FormatWith(FileCaseTemplateReferenceID));

                var startFileCases = new List<FileCase>();

                var fileCaseRecipientGroups =
                    f.GetProperty("#LHIND_Mitbericht_Weiterleiten") as SpecialdataListPropertyValueCollection;

                if (fileCaseRecipientGroups == null)
                    throw new ActivityCommandException("")
                        .WithUserMessage(WorkflowLocalization.NoRecipients.ToLocalizedName().FormatWith(f));

                foreach (var recipientGrp in fileCaseRecipientGroups)
                {
                    var newFC = FileCase.NewObject(f, null, null, fileCaseTemplate_WL);

                    if (newFC == null)
                        throw new ActivityCommandException("")
                            .WithUserMessage(WorkflowLocalization.NoFileCase.ToLocalizedName().FormatWith(commandActivity));


                    newFC.LeadingGroup = recipientGrp.Unwrap() as TenantGroup;

                    startFileCases.Add(newFC);
                }

                ClientTransaction.Current.Commit();
                foreach (var startFileCase in startFileCases)
                    startFileCase.StartObject();
                ClientTransaction.Current.Commit();
            }
            catch (Exception ex)
            {
                ClientTransaction.Current.Rollback();
                s_logger.Error(ex.Message);
                throw;
            }
            return true;
        }


        public override string Validate(CommandActivity commandActivity)
        {
            return null;
        }

    }

    public class LhindMitberichtReturnToGS : ActivityCommandModule
    {
        private static readonly ILog s_logger = LogManager.GetLogger(typeof(LhindMitberichtReturnToGS));

        public LhindMitberichtReturnToGS()
            : base("LHIND_Mitbericht_Rueckmeldung_GS:ActivityCommandClassificationType")
        {
        }

        public override bool Execute(CommandActivity commandActivity)
        {
            var restoreCtxID = ApplicationContext.CurrentID;
            try
            {
                var f = commandActivity.WorkItem as File;

                if (f == null)
                    throw new ActivityCommandException("")
                        .WithUserMessage(WorkflowLocalization.NoFile.ToLocalizedName().FormatWith(commandActivity));

                if (!f.CanEdit(true))
                    throw new ActivityCommandException("")
                        .WithUserMessage(WorkflowLocalization.NotEditable.ToLocalizedName().FormatWith(f));

                //Parse information from incoming
                var fcURL = String.Empty;

                //Get URL from incoming
                fcURL = f.BaseIncomings.FirstOrDefault().GetProperty("#LHIND_Mitbericht_VE_BaseObjectURL") as String;

                Uri fcUri = new Uri(fcURL);

                var tenantId = HttpUtility.ParseQueryString(fcUri.Query).Get("TenantID");
                var fcReferenceId = HttpUtility.ParseQueryString(fcUri.Query).Get("ObjectToOpenID").Substring(10, 36);
                //s_logger.Fatal(tenantId);
                //s_logger.Fatal(fcReferenceId);
                var fileContentObjects = f.GetFileContentHierarchyFlat().OfType<FileContentObject>().Where(c => c.ApprovalState == ApprovalStateType.Values.Approved() && c.HasReadAccess()).ToList();
                s_logger.Fatal(fileContentObjects);
                
                using (ClientTransaction.CreateRootTransaction().EnterDiscardingScope())
                using (TenantSection.SwitchToTenant(Tenant.FindByUnqiueIdentifier(tenantId))) 
                {
                    var targetFileCase = new ReferenceHandle<FileCase>(fcReferenceId).GetObject();
                    var incoming = Incoming.NewObject(targetFileCase.BaseFile);
                    ApplicationContext.CurrentID = incoming.ApplicationContextID;
                    incoming.LeadingGroup = UserHelper.Current.GetActaNovaUserExtension().StandardGroup != null
                        ? UserHelper.Current.GetActaNovaUserExtension().StandardGroup.AsTenantGroup()
                        : UserHelper.Current.OwningGroup.AsTenantGroup();

                    incoming.Subject = f.DisplayNameInternal;

                    //Eingangstyp und Fachdaten
                    var incomingTypeReferenceID = "64E46ED4-14C2-4E3D-A5D0-F3E2E39C2E73";
                    incoming.IncomingType = new ReferenceHandle<IncomingTypeClassificationType>(incomingTypeReferenceID).GetObject();

                    incoming.ExternalNumber = f.FormattedNumber;
                    incoming.Remark = f.Remark;
                    var template = Template.GetTemplateByNameOrGuid("Eingang: Erstellen");
                    ((IStartMainProcessDomainObjectMixin)incoming).CreateMainProcessInstance(template);
                    
                    foreach (var fileContentObject in fileContentObjects)
                    {
                        var copyDocument = (IDocument)Document.NewObject(incoming);
                        var originalDocument = fileContentObject.PrimaryContent;
                        copyDocument.Name = originalDocument.Name;
                        using(SecurityFreeSection.Activate())
                        using (TenantSection.DisableQueryRestrictions())
                        {
                            if (originalDocument.HasContent())
                            {
                                using (var handle = originalDocument.GetContent())
                                    copyDocument.SetContent(handle, originalDocument.Extension, originalDocument.MimeType);
                            }
                            else
                            {
                                copyDocument.MimeType = originalDocument.MimeType;
                                copyDocument.Extension = originalDocument.Extension;
                            }
                        }
                    }

                    targetFileCase.AddFileCaseContent(incoming);
                    
                    //File fGS = fcGS.ParentObject as File;
                    //fGS.add

                    ClientTransaction.Current.Commit();
                }
            }
            catch (Exception ex)
            {
                s_logger.Error(ex.Message);
                throw;
            }
            finally
            {
                ApplicationContext.CurrentID = restoreCtxID;
            }
            return true;
        }


        public override string Validate(CommandActivity commandActivity)
        {
            return null;
        }

    }

}