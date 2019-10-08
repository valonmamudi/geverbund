// This file is part of Acta Nova (www.acta-nova.eu)
// Copyright (c) rubicon IT GmbH, www.rubicon.eu
// Version 1.6 - Philipp Rössler - 03.10.2019

using System;
using System.Collections.Generic;
using System.Globalization;
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
using CsQuery.ExtensionMethods.Internal;
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
using Rubicon.Gever.Bund.Domain.Extensions;
using Rubicon.Gever.Bund.Domain.GeverService;
using Rubicon.Gever.Bund.Domain.Utilities;
using Rubicon.Gever.Bund.Domain.Utilities.Extensions;
using Rubicon.Multilingual.Extensions;
using Rubicon.ObjectBinding.Utilities;
using Rubicon.Utilities;
using Rubicon.Utilities.Globalization;
using Rubicon.Utilities.Security;
using Rubicon.Workflow.Domain;
using BusinessObjectExtensions = Remotion.ObjectBinding.BusinessObjectExtensions;
using Document = ActaNova.Domain.Document;

namespace LHIND.Mitbericht
{
    [LocalizationEnum]
    public enum LocalizedUserMessages
    {
        [MultiLingualName("Das Geschäftsobjekt der Aktivität \"{0}\" muss ein Dossier sein.", ""),
         MultiLingualName("L’objet métier de l’activité \"{0}\" doit être un dossier.","Fr"),
         MultiLingualName("L'oggetto business dell'attività \"{0}\" deve essere un dossier.", "It")]
        NoFile,

        [MultiLingualName("Das Geschäftsobjekt der Aktivität \"{0}\" muss ein Geschäftsvorfall sein.", ""),
         MultiLingualName("L’objet métier de l’activité \"{0}\" doit être une opération d’affaire.","Fr"),
         MultiLingualName("L'oggetto business dell'attività \"{0}\" deve essere un'operazione business.", "It")]
        NoFileCase,

        [MultiLingualName("Das Geschäftsobjekt \"{0}\" der Aktivität kann nicht bearbeitet werden.", ""),
         MultiLingualName("L’objet métier \"{0}\" de l’activité ne peut être édité.","Fr"),
         MultiLingualName("L'oggetto business \"{0}\" dell'attività non può essere modificato.", "It")]
        NotEditable,

        [MultiLingualName("Es wurde keine Geschäftsobjektvorlage mit der Referenz-ID \"{0}\" gefunden.", ""),
         MultiLingualName("Aucun modèle d’objets métier avec le numéro de référence \"{0}\" n’a été trouvé.","Fr"),
         MultiLingualName("Non è stato trovato alcun modello di oggetto business con l'ID di riferimento.", "It")]
        NoFileCaseTemplate,

        [MultiLingualName("Es wurde noch keine Federführung Amt definiert.", ""),
         MultiLingualName("Aucune unité d’organisation responsable n’a été définie.","Fr"),
         MultiLingualName("Non è stato ancora definito un ufficio responsabile.", "It")]
        NoFfDefined,

        [MultiLingualName("Es wurde keine Gruppe für den Katalogwert \"{0}\" gefunden.", ""),
         MultiLingualName("Es wurde keine Gruppe fär den Katalogwert \"{0}\" gefunden.","Fr"),
         MultiLingualName("Non è stato trovato nessun gruppo per il valore del catalogo \"{0}\".", "It")]
        NoGroupDefined,

        [MultiLingualName("Es wurde keine Geschäftsart \"{0}\" gefunden.", ""),
         MultiLingualName("Aucun groupe pour la valeur de catalogue \"{0}\" n’a été trouvé.","Fr"),
         MultiLingualName("Non è stato trovato nessun gruppo per il valore del catalogo \"{0}\".", "It")]
        NoFileCaseType,

        [MultiLingualName("Es wurden keine Empfänger \"{0}\" gefunden.", ""),
         MultiLingualName("Aucun type d’affaire \"{0}\" n’a été trouvé.","Fr"),
         MultiLingualName("Non è stato trovato nessun destinatario \"{0}\".", "It")]
        NoRecipients
    }

    public class LhindMitberichtCreateFilecasesAndAssignActivityCommandModules : ActivityCommandModule
    {
        private static readonly ILog logger =
            LogManager.GetLogger(typeof(LhindMitberichtCreateFilecasesAndAssignActivityCommandModules));

        public LhindMitberichtCreateFilecasesAndAssignActivityCommandModules() : base(
            "LHIND_Mitbericht_GVF_GS_erzeugen:ActivityCommandClassificationType")
        {
        }

        public override bool Execute(CommandActivity commandActivity)
        {
            //Set templates for FileCase creation.
            var fileCaseFfTemplateReferenceId = "6220CB49-8E09-4AB7-B4F1-673C6C91CC7E";
            var fileCaseMbTemplateReferenceId = "56E01B36-E6DE-4DA3-ABE8-95C6551A76D3";

            var startFileCases = new List<FileCase>();

            try
            {
                var f = commandActivity.WorkItem as File;

                if (f == null)
                    throw new ActivityCommandException("").WithUserMessage(LocalizedUserMessages.NoFile
                        .ToLocalizedName()
                        .FormatWith(commandActivity));

                if (!f.CanEdit(true))
                    throw new ActivityCommandException("").WithUserMessage(LocalizedUserMessages.NotEditable
                        .ToLocalizedName()
                        .FormatWith(f));

                //Create FF FileCase
                var fileCaseTemplateFf =
                    new ReferenceHandle<FileCaseTemplate>(fileCaseFfTemplateReferenceId).GetObject();
                if (fileCaseTemplateFf == null)
                    throw new ActivityCommandException("").WithUserMessage(LocalizedUserMessages.NoFileCaseTemplate
                        .ToLocalizedName()
                        .FormatWith(fileCaseFfTemplateReferenceId));

                var ffCatalogValue = f.GetProperty("#LHIND_Mitbericht_federfuhrendesAmt") as SpecialdataCatalogValue;
                if (ffCatalogValue == null)
                    throw new ActivityCommandException("").WithUserMessage(LocalizedUserMessages.NoFfDefined
                        .ToLocalizedName()
                        .FormatWith());

                var ffGroup = ffCatalogValue.GetProperty("#LHIND_Mitbericht_SPOC") as TenantGroup;
                if (ffGroup == null)
                    throw new ActivityCommandException("").WithUserMessage(LocalizedUserMessages.NoGroupDefined
                        .ToLocalizedName()
                        .FormatWith(ffCatalogValue));

                var ffFileCase = FileCase.NewObject(f, null, null, fileCaseTemplateFf);
                ffFileCase.LeadingGroup = ffGroup;

                var newTitle = ffFileCase.GetMultilingualValue(fc => fc.Title) + " " +
                               ffGroup.GetMultilingualValue(g => g.ShortName);
                ffFileCase.SetMultilingualValue(fc => fc.Title, newTitle);

                startFileCases.Add(ffFileCase);

                //Create MB FileCases
                var fileCaseTemplateMb =
                    new ReferenceHandle<FileCaseTemplate>(fileCaseMbTemplateReferenceId).GetObject();
                if (fileCaseTemplateMb == null)
                    throw new ActivityCommandException("").WithUserMessage(LocalizedUserMessages.NoFileCaseTemplate
                        .ToLocalizedName()
                        .FormatWith(fileCaseMbTemplateReferenceId));

                var mbCatalogValues =
                    f.GetProperty("#LHIND_Mitbericht_mitbeteiligtFDListe") as SpecialdataListPropertyValueCollection;

                foreach (var mbCatalogValue in mbCatalogValues)
                {
                    var mbGroup = mbCatalogValue.WrappedValue.GetProperty("#LHIND_Mitbericht_SPOC") as TenantGroup;
                    if (mbGroup == null)
                        throw new ActivityCommandException("").WithUserMessage(LocalizedUserMessages.NoGroupDefined
                            .ToLocalizedName()
                            .FormatWith(ffCatalogValue));

                    var mbFileCase = FileCase.NewObject(f, null, null, fileCaseTemplateMb);
                    mbFileCase.LeadingGroup = mbGroup;

                    newTitle = mbFileCase.GetMultilingualValue(fc => fc.Title) + " " +
                               mbGroup.GetMultilingualValue(g => g.ShortName);
                    mbFileCase.SetMultilingualValue(fc => fc.Title, newTitle);

                    startFileCases.Add(mbFileCase);
                }

                ClientTransaction.Current.Commit();

                foreach (var startFileCase in startFileCases) startFileCase.StartObject();

                ClientTransaction.Current.Commit();
            }
            catch (Exception ex)
            {
                logger.Error(ex.Message);
                ClientTransaction.Current.Rollback();
                throw;
            }

            return true;
        }

        public override string Validate(CommandActivity commandActivity)
        {
            return null;
        }
    }
    //End of LhindMitberichtCreateFilecasesAndAssignActivityCommandModules

    public class LhindMitberichtCreateNewIncomingFromFilecaseActivityCommandModules : ActivityCommandModule
    {
        private static readonly ILog logger =
            LogManager.GetLogger(typeof(LhindMitberichtCreateNewIncomingFromFilecaseActivityCommandModules));

        public LhindMitberichtCreateNewIncomingFromFilecaseActivityCommandModules() : base(
            "LHIND_Mitbericht_Eingang_erzeugen:ActivityCommandClassificationType")
        {
        }

        public override bool Execute(CommandActivity commandActivity)
        {
            //Set template reference for FF
            var fileCaseFfTypeReferenceId = "589AFF8A-7737-40DE-BBA4-8FA2A3178058";
            var incomingTypeVeReferenceId = "ABC5B6D0-6764-4801-AC36-7B42B80F63D3";

            var restoreCtxId = ApplicationContext.CurrentID;

            try
            {
                var fc = commandActivity.WorkItem as FileCase;

                if (fc == null)
                    throw new ActivityCommandException("").WithUserMessage(LocalizedUserMessages.NoFileCase
                        .ToLocalizedName()
                        .FormatWith(commandActivity));

                if (!fc.CanEdit(true))
                    throw new ActivityCommandException("").WithUserMessage(LocalizedUserMessages.NotEditable
                        .ToLocalizedName()
                        .FormatWith(fc));

                //Read specialdata and document values
                var terminAmtObject = fc.GetProperty("#LHIND_Mitbericht_BezeichnerDU2").ToString();
                var terminAmt = "";
                if (terminAmtObject != null)
                    terminAmt = terminAmtObject;

                var terminGsObject = fc.GetProperty("#LHIND_Mitbericht_BezeichnerDU1").ToString();
                var terminGs = "";
                if (terminGsObject != null)
                    terminGs = terminGsObject;

                var ffAmtObject = fc.GetProperty("#LHIND_Mitbericht_federfuhrendesAmt") as SpecialdataCatalogValue;
                var ffAmt = "";
                if (ffAmtObject != null)
                    ffAmt = ffAmtObject.DisplayName;

                var rueckfragenAnObject = fc.GetProperty("#LHIND_Mitbericht_rückfragenAn") as TenantUser;
                var rueckfragenAn = "";
                if (rueckfragenAnObject != null)
                    rueckfragenAn = rueckfragenAnObject.DisplayName;

                var mbAemterCatalogValues =
                    fc.GetProperty("#LHIND_Mitbericht_mitbeteiligtFDListe") as
                        SpecialdataListPropertyValueCollection;
                var mbAemter = String.Join(", ",
                    mbAemterCatalogValues.ToList()
                        .Select(m => m.WrappedValue.GetProperty("DisplayName"))
                        .ToList()
                        .ToArray());

                var title = fc.GetProperty("#LHIND_Mitbericht_titelIDP") as string;

                var auftragsartObject =
                    fc.GetProperty("#LHIND_Mitbericht_auftragsartMitberichtsverfahren") as SpecialdataCatalogValue;
                var auftragsart = "";
                if (auftragsartObject != null)
                    auftragsart = auftragsartObject.DisplayName;

                var rueckmeldungCatalogValues =
                    fc.GetProperty("#LHIND_Mitbericht_rückmeldungAn") as SpecialdataListPropertyValueCollection;
                var rueckmeldung = String.Join(", ",
                    rueckmeldungCatalogValues.ToList()
                        .Select(m => m.WrappedValue.GetProperty("DisplayName"))
                        .ToList()
                        .ToArray());

                var fcUrl = UrlProvider.Current.GetOpenWorkListItemUrl(fc);
                bool isFf = (fc.Type.ToHasReferenceID().ReferenceID.ToUpper() == fileCaseFfTypeReferenceId.ToUpper());

                IList<FileContentObject> fileContentObjects;
                using (TenantSection.DisableQueryRestrictions())
                    fileContentObjects = fc.FileCaseContents
                        .Where(c => c.Document != null && c.Document.ActiveContent.HasContent())
                        .Select(c => c.Document)
                        .ToList();

                //switch tenant
                using (ClientTransaction.CreateRootTransaction().EnterDiscardingScope())
                using (TenantSection.SwitchToTenant(UserHelper.Current.Tenant))
                {
                    //Create new incoming, set specialdata properties
                    var incoming = Incoming.NewObject();
                    ApplicationContext.CurrentID = incoming.ApplicationContextID;
                    incoming.LeadingGroup = UserHelper.Current.GetActaNovaUserExtension().StandardGroup != null
                        ? UserHelper.Current.GetActaNovaUserExtension().StandardGroup.AsTenantGroup()
                        : UserHelper.Current.OwningGroup.AsTenantGroup();

                    incoming.Subject = fc.DisplayName;
                    incoming.IncomingType =
                        new ReferenceHandle<IncomingTypeClassificationType>(incomingTypeVeReferenceId).GetObject();
                    incoming.ExternalNumber = fc.FormattedNumber;
                    incoming.Remark = fc.WorkInstruction;

                    using (new SpecialdataIgnoreReadOnlySection())
                    {
                        incoming.SetProperty("#LHIND_Mitbericht_VE_TerminGS", terminGs);
                        incoming.SetProperty("#LHIND_Mitbericht_VE_TerminAmt", terminAmt);
                        incoming.SetProperty("#LHIND_Mitbericht_VE_Rückmeldung", rueckmeldung);
                        incoming.SetProperty("#LHIND_Mitbericht_VE_ffAmt", ffAmt);
                        incoming.SetProperty("#LHIND_Mitbericht_VE_AuftragsartMitbericht", auftragsart);
                        incoming.SetProperty("#LHIND_Mitbericht_VE_Titel", title);
                        incoming.SetProperty("#LHIND_Mitbericht_VE_Mitbeteiligt", mbAemter);
                        incoming.SetProperty("#LHIND_Mitbericht_VE_RückfragenAn", rueckfragenAn);
                        incoming.SetProperty("#LHIND_Mitbericht_VE_BaseObjectURL", fcUrl);
                        incoming.SetProperty("#LHIND_Mitbericht_VE_istFederfuehrung", isFf);
                    }

                    foreach (var fileContentObject in fileContentObjects)
                    {
                        var copyDocument = (IDocument)Document.NewObject(incoming);
                        var originalDocument = fileContentObject.ActiveContent;

                        copyDocument.Name = originalDocument.Name;

                        using (TenantSection.DisableQueryRestrictions())
                        using (var handle = originalDocument.GetContent())
                            copyDocument.SetContent(handle, originalDocument.Extension, originalDocument.MimeType);
                        
                        ClientTransaction.Current.Commit();
                    }
                }
            }
            catch (Exception ex)
            {
                logger.Error(ex.Message);
                throw;
            }
            finally
            {
                ApplicationContext.CurrentID = restoreCtxId;
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
        private static readonly ILog logger = LogManager.GetLogger(typeof(LhindMitberichtCreateInternalFilecases));

        public LhindMitberichtCreateInternalFilecases() : base(
            "LHIND_Mitbericht_Interne_Weiterleitung:ActivityCommandClassificationType")
        {
        }

        public override bool Execute(CommandActivity commandActivity)
        {
            var fileCaseTemplateReferenceId = "04248329-15C8-40C8-8888-90DF1C782A56";

            try
            {
                var f = commandActivity.WorkItem as File;

                if (f == null)
                    throw new ActivityCommandException("").WithUserMessage(LocalizedUserMessages.NoFile
                        .ToLocalizedName()
                        .FormatWith(commandActivity));

                if (!f.CanEdit(true))
                    throw new ActivityCommandException("").WithUserMessage(LocalizedUserMessages.NotEditable
                        .ToLocalizedName()
                        .FormatWith(f));

                var fileCaseTemplate = new ReferenceHandle<FileCaseTemplate>(fileCaseTemplateReferenceId).GetObject();

                if (fileCaseTemplate == null)
                    throw new ActivityCommandException("").WithUserMessage(LocalizedUserMessages.NoFileCaseTemplate
                        .ToLocalizedName()
                        .FormatWith(fileCaseTemplateReferenceId));

                var startFileCases = new List<FileCase>();

                var fileCaseRecipientGroups =
                    f.GetProperty("#LHIND_Mitbericht_Weiterleiten") as SpecialdataListPropertyValueCollection;

                if (fileCaseRecipientGroups == null)
                    throw new ActivityCommandException("").WithUserMessage(LocalizedUserMessages.NoRecipients
                        .ToLocalizedName()
                        .FormatWith(f));

                foreach (var recipientGroup in fileCaseRecipientGroups)
                {
                    var newFileCase = FileCase.NewObject(f, null, null, fileCaseTemplate);

                    if (newFileCase == null)
                        throw new ActivityCommandException("").WithUserMessage(LocalizedUserMessages.NoFileCase
                            .ToLocalizedName()
                            .FormatWith(commandActivity));

                    newFileCase.LeadingGroup = recipientGroup.Unwrap() as TenantGroup;

                    startFileCases.Add(newFileCase);
                }

                ClientTransaction.Current.Commit();
                foreach (var startFileCase in startFileCases) startFileCase.StartObject();
                ClientTransaction.Current.Commit();
            }
            catch (Exception ex)
            {
                logger.Error(ex.Message);
                ClientTransaction.Current.Rollback();
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
        private static readonly ILog logger = LogManager.GetLogger(typeof(LhindMitberichtReturnToGS));

        public LhindMitberichtReturnToGS() : base("LHIND_Mitbericht_Rueckmeldung_GS:ActivityCommandClassificationType")
        {
        }

        public override bool Execute(CommandActivity commandActivity)
        {
            var incomingTypeVeReferenceId = "ABC5B6D0-6764-4801-AC36-7B42B80F63D3";
            var incomingTypeGsReferenceId = "64E46ED4-14C2-4E3D-A5D0-F3E2E39C2E73";

            var restoreCtxId = ApplicationContext.CurrentID;

            try
            {
                var f = commandActivity.WorkItem as File;

                if (f == null)
                    throw new ActivityCommandException("").WithUserMessage(LocalizedUserMessages.NoFile
                        .ToLocalizedName()
                        .FormatWith(commandActivity));

                if (!f.CanEdit(true))
                    throw new ActivityCommandException("").WithUserMessage(LocalizedUserMessages.NotEditable
                        .ToLocalizedName()
                        .FormatWith(f));

                //Parse information from incoming
                var fcUri = new Uri(f.BaseIncomings
                    .First(i => i.IncomingType.ToHasReferenceID().ReferenceID.ToUpper() ==
                                incomingTypeVeReferenceId.ToUpper())
                    .GetProperty("#LHIND_Mitbericht_VE_BaseObjectURL") as string);
                var tenantId = HttpUtility.ParseQueryString(fcUri.Query).Get("TenantID");
                var fcReferenceId = HttpUtility.ParseQueryString(fcUri.Query).Get("ObjectToOpenID").Substring(10, 36);

                IList<FileContentObject> fileContentObjects;
                using (TenantSection.DisableQueryRestrictions())
                    fileContentObjects = f.GetFileContentHierarchyFlat()
                        .OfType<FileContentObject>()
                        .Where(c => c.ApprovalState == ApprovalStateType.Values.Approved() && c.HasReadAccess() &&
                                    c.ActiveContent.HasContent())
                        .ToList();

                using (ClientTransaction.CreateRootTransaction().EnterDiscardingScope())
                using (TenantSection.SwitchToTenant(Tenant.FindByUnqiueIdentifier(tenantId)))
                {
                    var targetFileCase = new ReferenceHandle<FileCase>(fcReferenceId).GetObject();
                    var incoming = Incoming.NewObject(targetFileCase.BaseFile);
                    ApplicationContext.CurrentID = incoming.ApplicationContextID;
                    incoming.LeadingGroup = UserHelper.Current.GetActaNovaUserExtension().StandardGroup != null
                        ? UserHelper.Current.GetActaNovaUserExtension().StandardGroup.AsTenantGroup()
                        : UserHelper.Current.OwningGroup.AsTenantGroup();

                    incoming.Subject = f.DisplayName;
                    incoming.IncomingType =
                        new ReferenceHandle<IncomingTypeClassificationType>(incomingTypeGsReferenceId).GetObject();
                    incoming.ExternalNumber = f.FormattedNumber;
                    incoming.Remark = f.Remark;

                    foreach (var fileContentObject in fileContentObjects)
                    {
                        var originalDocument = fileContentObject.ActiveContent;
                        var copyDocument = (IDocument) Document.NewObject(incoming);

                        copyDocument.Name = originalDocument.Name;

                        using (SecurityFreeSection.Activate())
                        using (TenantSection.DisableQueryRestrictions())
                        using (var handle = originalDocument.GetContent())
                            copyDocument.SetContent(handle, originalDocument.Extension, originalDocument.MimeType);
                    }

                    targetFileCase.AddFileCaseContent(incoming);

                    ClientTransaction.Current.Commit();
                }
            }
            catch (Exception ex)
            {
                logger.Error(ex.Message);
                throw;
            }
            finally
            {
                ApplicationContext.CurrentID = restoreCtxId;
            }

            return true;
        }

        public override string Validate(CommandActivity commandActivity)
        {
            return null;
        }
    }
}