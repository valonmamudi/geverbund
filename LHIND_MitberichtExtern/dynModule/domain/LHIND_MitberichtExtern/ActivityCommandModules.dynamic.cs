// This file is part of Acta Nova (www.acta-nova.eu)
// Copyright (c) rubicon IT GmbH, www.rubicon.eu
// Version 2.2 - Philipp Rössler, Mohammad-Farid Modarressi - 28.10.2019

using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using ActaNova.Domain;
using ActaNova.Domain.Classes.Configuration;
using ActaNova.Domain.Classifications;
using ActaNova.Domain.Extensions;
using ActaNova.Domain.Specialdata;
using ActaNova.Domain.Specialdata.Catalog;
using ActaNova.Domain.Specialdata.Values;
using ActaNova.Domain.Workflow;
using Remotion.Data.DomainObjects;
using Remotion.Globalization;
using Remotion.Logging;
using Remotion.ObjectBinding;
using Remotion.Security;
using Remotion.SecurityManager.Domain.OrganizationalStructure;
using Rubicon.Dms;
using Rubicon.Domain;
using Rubicon.Domain.TenantBoundOrganizationalStructure;
using Rubicon.Gever.Bund.Domain.Utilities.Extensions;
using Rubicon.Gever.Bund.EGovInterface.Domain;
using Rubicon.Gever.Bund.EGovInterface.Domain.Export;
using Rubicon.Gever.Bund.EGovInterface.Domain.Import;
using Rubicon.Gever.Bund.EGovInterface.Domain.Import.Mapping;
using Rubicon.Gever.Bund.EGovInterface.Domain.Reader;
using Rubicon.Multilingual.Extensions;
using Rubicon.ObjectBinding.Utilities;
using Rubicon.Utilities;
using Rubicon.Utilities.Autofac;
using Rubicon.Utilities.Globalization;
using Rubicon.Utilities.Security;
using Rubicon.Workflow.Domain;

namespace LHIND.MitberichtExtern
{
    [LocalizationEnum]
    public enum LocalizedUserMessages
    {
        [MultiLingualName("Das Geschäftsobjekt der Aktivität \"{0}\" muss ein Dossier sein.", "")]
        [MultiLingualName("L’objet métier de l’activité \"{0}\" doit être un dossier.", "Fr")]
        [MultiLingualName("L'oggetto business dell'attività \"{0}\" deve essere un dossier.", "It")]
        NoFile,

        [MultiLingualName("Das Geschäftsobjekt der Aktivität \"{0}\" muss ein Geschäftsvorfall sein.", "")]
        [MultiLingualName("L’objet métier de l’activité \"{0}\" doit être une opération d’affaire.", "Fr")]
        [MultiLingualName("L'oggetto business dell'attività \"{0}\" deve essere un'operazione business.", "It")]
        NoFileCase,

        [MultiLingualName("Das Geschäftsobjekt \"{0}\" der Aktivität kann nicht bearbeitet werden.", "")]
        [MultiLingualName("L’objet métier \"{0}\" de l’activité ne peut être édité.", "Fr")]
        [MultiLingualName("L'oggetto business \"{0}\" dell'attività non può essere modificato.", "It")]
        NotEditable,

        [MultiLingualName("Es wurde keine Geschäftsobjektvorlage mit der Referenz-ID \"{0}\" gefunden.", "")]
        [MultiLingualName("Aucun modèle d’objets métier avec le numéro de référence \"{0}\" n’a été trouvé.", "Fr")]
        [MultiLingualName("Non è stato trovato alcun modello di oggetto business con l'ID di riferimento.", "It")]
        NoFileCaseTemplate,

        [MultiLingualName("Es wurde noch keine Federführung Amt definiert.", "")]
        [MultiLingualName("Aucune unité d’organisation responsable n’a été définie.", "Fr")]
        [MultiLingualName("Non è stato ancora definito un ufficio responsabile.", "It")]
        NoFfDefined,

        [MultiLingualName("Es wurde keine Gruppe für den Katalogwert \"{0}\" gefunden.", "")]
        [MultiLingualName("Es wurde keine Gruppe fär den Katalogwert \"{0}\" gefunden.", "Fr")]
        [MultiLingualName("Non è stato trovato nessun gruppo per il valore del catalogo \"{0}\".", "It")]
        NoGroupDefined,

        [MultiLingualName("Es wurde keine Geschäftsart \"{0}\" gefunden.", "")]
        [MultiLingualName("Aucun groupe pour la valeur de catalogue \"{0}\" n’a été trouvé.", "Fr")]
        [MultiLingualName("Non è stato trovato nessun gruppo per il valore del catalogo \"{0}\".", "It")]
        NoFileCaseType,

        [MultiLingualName("Es wurden keine Empfänger \"{0}\" gefunden.", "")]
        [MultiLingualName("Aucun type d’affaire \"{0}\" n’a été trouvé.", "Fr")]
        [MultiLingualName("Non è stato trovato nessun destinatario \"{0}\".", "It")]
        NoRecipients,

        [MultiLingualName("Das eCH-Paket wurde noch nicht importiert.", "")]
        [MultiLingualName("Le paquet n'a pas encore été importé.", "Fr")]
        [MultiLingualName("Il pacco non è stato ancora consegnato.", "It")]
        FileNotImported,

        [MultiLingualName("Es konnte kein Eingang mit Referenz zum Quellobjekt im Dossier (\"{0}\") gefunden werden.",
            "")]
        [MultiLingualName("...", "Fr")]
        [MultiLingualName("...", "It")]
        IncomingNotFound
    }

    public class LhindMitberichtExternCreateIdpFilecases : ActivityCommandModule
    {
        private static readonly ILog s_logger =
            LogManager.GetLogger(typeof(LhindMitberichtExternCreateIdpFilecases));

        public LhindMitberichtExternCreateIdpFilecases() : base(
            "LHIND_MitberichtExtern_IdpGvfErzeugen:ActivityCommandClassificationType")
        {
        }

        public override bool Execute(CommandActivity commandActivity)
        {
            //Set templates for FileCase creation.
            var fileCaseFfTemplateReferenceId = "B475F10A-94ED-4E75-9C64-DF6890569093";
            var fileCaseMbTemplateReferenceId = "8087873D-755F-40D4-84C6-1BA7EB8C26AC";

            var startFileCases = new List<FileCase>();

            try
            {
                var hostObject = (File)commandActivity.WorkItem;

                if (hostObject == null)
                    throw new ActivityCommandException("").WithUserMessage(LocalizedUserMessages.NoFile
                        .ToLocalizedName()
                        .FormatWith(commandActivity));

                if (!hostObject.CanEdit(true))
                    throw new ActivityCommandException("").WithUserMessage(LocalizedUserMessages.NotEditable
                        .ToLocalizedName()
                        .FormatWith(hostObject));

                //Create FF FileCase
                var fileCaseTemplateFf =
                    new ReferenceHandle<FileCaseTemplate>(fileCaseFfTemplateReferenceId).GetObject();
                if (fileCaseTemplateFf == null)
                    throw new ActivityCommandException("").WithUserMessage(LocalizedUserMessages.NoFileCaseTemplate
                        .ToLocalizedName()
                        .FormatWith(fileCaseFfTemplateReferenceId));
                var newTitle = String.Empty;

                var ffCatalogValue = hostObject.GetProperty("#LHIND_MitberichtExtern_FederfuhrendesAmt") as SpecialdataCatalogValue;
                if (ffCatalogValue != null)
                {
                    var ffGroup = ffCatalogValue.GetProperty("#LHIND_MitberichtExtern_Spoc") as TenantGroup;
                    if (ffGroup == null)
                        throw new ActivityCommandException("").WithUserMessage(LocalizedUserMessages.NoGroupDefined
                            .ToLocalizedName()
                            .FormatWith(ffCatalogValue));

                    var ffFileCase = FileCase.NewObject(hostObject, null, null, fileCaseTemplateFf);
                    ffFileCase.LeadingGroup = ffGroup;

                    newTitle = ffFileCase.GetMultilingualValue(fc => fc.Title) + " - " + ffGroup.GetMultilingualValue(g => g.ShortName);
                    ffFileCase.SetMultilingualValue(fc => fc.Title, newTitle);

                    startFileCases.Add(ffFileCase);
                }

                //Create MB FileCases
                var fileCaseTemplateMb =
                    new ReferenceHandle<FileCaseTemplate>(fileCaseMbTemplateReferenceId).GetObject();
                if (fileCaseTemplateMb == null)
                    throw new ActivityCommandException("").WithUserMessage(LocalizedUserMessages.NoFileCaseTemplate
                        .ToLocalizedName()
                        .FormatWith(fileCaseMbTemplateReferenceId));

                var mbCatalogValues =
                    hostObject.GetProperty("#LHIND_MitberichtExtern_MitbeteiligtFdListe") as SpecialdataListPropertyValueCollection;

                foreach (var mbCatalogValue in mbCatalogValues)
                {
                    var mbGroup = mbCatalogValue.WrappedValue.GetProperty("#LHIND_MitberichtExtern_Spoc") as TenantGroup;
                    if (mbGroup == null)
                        throw new ActivityCommandException("").WithUserMessage(LocalizedUserMessages.NoGroupDefined
                            .ToLocalizedName()
                            .FormatWith(ffCatalogValue));

                    var mbFileCase = FileCase.NewObject(hostObject, null, null, fileCaseTemplateMb);
                    mbFileCase.LeadingGroup = mbGroup;

                    newTitle = mbFileCase.GetMultilingualValue(fc => fc.Title) + " - " + mbFileCase.LeadingGroup.GetMultilingualValue(g => g.ShortName);
                    mbFileCase.SetMultilingualValue(fc => fc.Title, newTitle);

                    startFileCases.Add(mbFileCase);
                }

                ClientTransaction.Current.Commit();

                foreach (var startFileCase in startFileCases) startFileCase.StartObject();

                ClientTransaction.Current.Commit();
            }
            catch (Exception ex)
            {
                s_logger.Error(ex.Message, ex);
                ClientTransaction.Current.Rollback();
                throw;
            }

            return true;
        }

        public override string Validate(CommandActivity commandActivity)
        {
            if (commandActivity.EffectiveWorkItemType.IsAssignableFrom(typeof(File)))
                return null;
            else
                return "Invalid WorkItemType";
        }
    }
    //End of LhindMitberichtExternCreateFilecasesAndAssignActivityCommandModules

    public class LhindMitberichtExternCreateIncomingFromFilecase : ActivityCommandModule
    {
        private static readonly ILog s_logger =
            LogManager.GetLogger(typeof(LhindMitberichtExternCreateIncomingFromFilecase));

        public LhindMitberichtExternCreateIncomingFromFilecase() : base(
            "LHIND_MitberichtExtern_EingangErstellen:ActivityCommandClassificationType")
        {
        }

        public override bool Execute(CommandActivity commandActivity)
        {
            //Set template reference for FF
            var fileCaseFfTypeReferenceId = "A8DE3E44-A236-4CBA-91A0-AD16A2D734BA";
            var incomingTypeVeReferenceId = "3A655116-6B6E-4246-B2ED-7A213FD61493";

            var restoreCtxId = ApplicationContext.CurrentID;

            try
            {
                var sourceFileCase = (FileCase) commandActivity.WorkItem;

                if (sourceFileCase == null)
                    throw new ActivityCommandException("").WithUserMessage(LocalizedUserMessages.NoFileCase
                        .ToLocalizedName()
                        .FormatWith(commandActivity));

                if (!sourceFileCase.CanEdit(true))
                    throw new ActivityCommandException("").WithUserMessage(LocalizedUserMessages.NotEditable
                        .ToLocalizedName()
                        .FormatWith(sourceFileCase));

                //Read specialdata and document values
                var terminGsObject = sourceFileCase.GetProperty("#LHIND_MitberichtExtern_TerminGs").ToString();
                var terminGs = "";
                if (terminGsObject != null)
                    terminGs = terminGsObject;

                var datumObject = sourceFileCase.GetProperty("#LHIND_MitberichtExtern_Datum").ToString();
                var datum = "";
                if (datumObject != null)
                    datum = datumObject;

                var ffAmtObject =
                    sourceFileCase.GetProperty("#LHIND_MitberichtExtern_FederfuhrendesAmt") as SpecialdataCatalogValue;
                var ffAmt = "";
                if (ffAmtObject != null)
                    ffAmt = ffAmtObject.DisplayName;

                var rueckfragenAnObject =
                    sourceFileCase.GetProperty("#LHIND_MitberichtExtern_RückfragenAn") as TenantUser;
                var rueckfragenAn = "";
                if (rueckfragenAnObject != null)
                    rueckfragenAn = rueckfragenAnObject.DisplayName;

                var mbAemterCatalogValues =
                    sourceFileCase.GetProperty("#LHIND_MitberichtExtern_MitbeteiligtFdListe") as
                        SpecialdataListPropertyValueCollection;

                var mbAemter = string.Join(", ",
                    mbAemterCatalogValues.ToList()
                        .Select(m => m.WrappedValue.GetProperty("DisplayName"))
                        .ToList()
                        .ToArray());

                var title = sourceFileCase.GetProperty("#LHIND_MitberichtExtern_TitelIdp") as string;

                var auftragsartObject =
                    sourceFileCase.GetProperty("#LHIND_MitberichtExtern_AuftragsartMitberichtsverfahren") as
                        SpecialdataCatalogValue;
                var auftragsart = "";
                if (auftragsartObject != null)
                    auftragsart = auftragsartObject.DisplayName;

                var rueckmeldungCatalogValues =
                    sourceFileCase.GetProperty("#LHIND_MitberichtExtern_RückmeldungAn") as
                        SpecialdataListPropertyValueCollection;
                var rueckmeldung = string.Join(", ",
                    rueckmeldungCatalogValues.ToList()
                        .Select(m => m.WrappedValue.GetProperty("DisplayName"))
                        .ToList()
                        .ToArray());

                var bemerkungObject = sourceFileCase.GetProperty("#LHIND_MitberichtExtern_Bemerkungen") as string;
                var bemerkung = "";
                if (bemerkungObject != null)
                    bemerkung = bemerkungObject;

                var sourceFileCaseUrl = UrlProvider.Current.GetOpenWorkListItemUrl(sourceFileCase);
                bool istFederfuehrung = sourceFileCase.Type.ToHasReferenceID().ReferenceID.ToUpper() == fileCaseFfTypeReferenceId.ToUpper();

                //Create eCH-0147 container
                var messageExport = Containers.Global.Resolve<IMessageExport>();
                var eChExport = messageExport.Export(sourceFileCase);

                sourceFileCase.AddFileCaseContent(eChExport);

                //switch tenant
                using (ClientTransaction.CreateRootTransaction().EnterDiscardingScope())
                using (TenantSection.SwitchToTenant(UserHelper.Current.Tenant))
                {
                    //Create new incoming, set specialdata properties
                    var incoming = Incoming.NewObject();
                    ApplicationContext.CurrentID = incoming.ApplicationContextID;

                    incoming.Subject = sourceFileCase.DisplayName + " - eCH-Dossier";
                    incoming.IncomingType =
                        new ReferenceHandle<IncomingTypeClassificationType>(incomingTypeVeReferenceId).GetObject();
                    incoming.ExternalNumber = sourceFileCase.FormattedNumber;
                    incoming.Remark = sourceFileCase.WorkInstruction;

                    using (new SpecialdataIgnoreReadOnlySection())
                    {
                        incoming.SetProperty("#LHIND_MitberichtExtern_VE_TerminGs", terminGs);
                        incoming.SetProperty("#LHIND_MitberichtExtern_VE_Titel", title);
                        incoming.SetProperty("#LHIND_MitberichtExtern_VE_FfAmt", ffAmt);
                        incoming.SetProperty("#LHIND_MitberichtExtern_VE_Mitbeteiligt", mbAemter);
                        incoming.SetProperty("#LHIND_MitberichtExtern_VE_Rückmeldung", rueckmeldung);
                        incoming.SetProperty("#LHIND_MitberichtExtern_VE_AuftragsartMitbericht", auftragsart);
                        incoming.SetProperty("#LHIND_MitberichtExtern_VE_Bemerkungen", bemerkung);
                        incoming.SetProperty("#LHIND_MitberichtExtern_VE_Datum", datum);
                        incoming.SetProperty("#LHIND_MitberichtExtern_VE_RückfragenAn", rueckfragenAn);
                        incoming.SetProperty("#LHIND_MitberichtExtern_SourceFileCaseUrl", sourceFileCaseUrl);
                        incoming.SetProperty("#LHIND_MitberichtExtern_VE_IstFederfuehrung", istFederfuehrung);
                    }

                    var targeteChDocument = Document.NewObject(incoming);
                    ((IDocument) targeteChDocument).Name = sourceFileCase.GetMultilingualValue(fc => fc.Title) + " (" +
                                                           sourceFileCase.FormattedNumber + ") - eCH Import";
                    targeteChDocument.PhysicallyPresent = false;
                    targeteChDocument.Type =
                        (DocumentClassificationType) ClassificationType.GetObject(WellKnownObjects
                            .DocumentClassification.EchImport.GetObjectID());

                    using (TenantSection.DisableQueryRestrictions())
                    using (var handle = eChExport.ActiveContent.GetContent())
                    {
                        targeteChDocument.ActiveContent.SetContent(handle, "zip", "application/zip");
                    }

                    var targetFile = ImportHelper.TenantKnowsObject(targeteChDocument, true);
                    if (targetFile != null)
                    {
                        incoming.LeadingGroup = targetFile.LeadingGroup;
                        incoming.Insert(targetFile);
                    }
                    else
                    {
                        incoming.LeadingGroup = UserHelper.Current.GetActaNovaUserExtension().StandardGroup != null
                            ? UserHelper.Current.GetActaNovaUserExtension().StandardGroup.AsTenantGroup()
                            : UserHelper.Current.OwningGroup.AsTenantGroup();
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
                ApplicationContext.CurrentID = restoreCtxId;
            }

            return true;
        }

        public override string Validate(CommandActivity commandActivity)
        {
            if (commandActivity.EffectiveWorkItemType.IsAssignableFrom(typeof(FileCase)))
                return null;
            return "Invalid WorkItemType";
        }
    }
    //End of LhindMitberichtExternCreateIncomingFromFilecase


    public class LhindMitberichtExternRegisterIncomingToeCHImport : ActivityCommandModule
    {
        private static readonly ILog s_logger =
            LogManager.GetLogger(typeof(LhindMitberichtExternRegisterIncomingToeCHImport));

        public LhindMitberichtExternRegisterIncomingToeCHImport() : base(
            "LHIND_MitberichtExtern_eCHEingangRegistrieren:ActivityCommandClassificationType")
        {
        }

        public override bool Execute(CommandActivity commandActivity)
        {
            try
            {
                var hostObject = (Incoming) commandActivity.WorkItem;

                if (hostObject.BaseFile != null)
                    return true;

                //Get target file
                var echDocument = hostObject.FileContentHierarchyFlat.OfType<Document>()
                    .First(d => d.Type == ClassificationType.GetObject(WellKnownObjects.DocumentClassification.EchImport
                                    .GetObjectID()));

                var targetFile = ImportHelper.TenantKnowsObject(echDocument);

                if (targetFile == null)
                    throw new ActivityCommandException("").WithUserMessage(LocalizedUserMessages.FileNotImported
                        .ToLocalizedName()
                        .FormatWith(commandActivity));

                hostObject.Insert(targetFile);
            }
            catch (Exception ex)
            {
                s_logger.Error(ex.Message, ex);
                throw;
            }

            return true;
        }

        public override string Validate(CommandActivity commandActivity)
        {
            if (commandActivity.EffectiveWorkItemType.IsAssignableFrom(typeof(Incoming)))
                return null;
            return "Invalid WorkItemType";
        }
    }
    //End of LhindMitberichtExternRegisterIncomingToeCHImport


    public class LhindMitberichtExternCreateInternalFilecases : ActivityCommandModule
    {
        private static readonly ILog s_logger =
            LogManager.GetLogger(typeof(LhindMitberichtExternCreateInternalFilecases));

        public LhindMitberichtExternCreateInternalFilecases() : base(
            "LHIND_MitberichtExtern_InterneWeiterleitung:ActivityCommandClassificationType")
        {
        }

        public override bool Execute(CommandActivity commandActivity)
        {
            var fileCaseTemplateReferenceId = "0ADBF498-5C5E-4D7D-B1C3-39EA4DA65463";

            try
            {
                var hostObject = (File) commandActivity.WorkItem;

                if (hostObject == null)
                    throw new ActivityCommandException("").WithUserMessage(LocalizedUserMessages.NoFile
                        .ToLocalizedName()
                        .FormatWith(commandActivity));

                if (!hostObject.CanEdit(true))
                    throw new ActivityCommandException("").WithUserMessage(LocalizedUserMessages.NotEditable
                        .ToLocalizedName()
                        .FormatWith(hostObject));

                var fileCaseTemplate = new ReferenceHandle<FileCaseTemplate>(fileCaseTemplateReferenceId).GetObject();

                if (fileCaseTemplate == null)
                    throw new ActivityCommandException("").WithUserMessage(LocalizedUserMessages.NoFileCaseTemplate
                        .ToLocalizedName()
                        .FormatWith(fileCaseTemplateReferenceId));

                var startFileCases = new List<FileCase>();

                var fileCaseRecipientGroups =
                    hostObject.GetProperty("#LHIND_MitberichtExtern_Weiterleiten") as
                        SpecialdataListPropertyValueCollection;

                if (fileCaseRecipientGroups == null)
                    throw new ActivityCommandException("").WithUserMessage(LocalizedUserMessages.NoRecipients
                        .ToLocalizedName()
                        .FormatWith(hostObject));

                foreach (var recipientGroup in fileCaseRecipientGroups)
                {
                    var newFileCase = FileCase.NewObject(hostObject, null, null, fileCaseTemplate);

                    if (newFileCase == null)
                        throw new ActivityCommandException("").WithUserMessage(LocalizedUserMessages.NoFileCase
                            .ToLocalizedName()
                            .FormatWith(commandActivity));

                    newFileCase.LeadingGroup = recipientGroup.Unwrap() as TenantGroup;
                    var newTitle = hostObject.GetMultilingualValue(fc => fc.Title) + " [" + hostObject.FormattedNumber +
                                   "] - " + newFileCase.LeadingGroup.GetMultilingualValue(g => g.ShortName);
                    newFileCase.SetMultilingualValue(fc => fc.Title, newTitle);

                    startFileCases.Add(newFileCase);
                }

                ClientTransaction.Current.Commit();

                foreach (var startFileCase in startFileCases) startFileCase.StartObject();

                ClientTransaction.Current.Commit();
            }
            catch (Exception ex)
            {
                s_logger.Error(ex.Message);
                ClientTransaction.Current.Rollback();
                throw;
            }

            return true;
        }

        public override string Validate(CommandActivity commandActivity)
        {
            if (commandActivity.EffectiveWorkItemType.IsAssignableFrom(typeof(File)))
                return null;
            return "Invalid WorkItemType";
        }
    }

    public class LhindMitberichtExternReturnToGS : ActivityCommandModule
    {
        private static readonly ILog s_logger = LogManager.GetLogger(typeof(LhindMitberichtExternReturnToGS));

        public LhindMitberichtExternReturnToGS() : base(
            "LHIND_MitberichtExtern_RueckmeldungGs:ActivityCommandClassificationType")
        {
        }

        public override bool Execute(CommandActivity commandActivity)
        {
            var incomingTypeVeReferenceId = "3A655116-6B6E-4246-B2ED-7A213FD61493";
            var incomingTypeGsReferenceId = "2C22B34E-29D6-44AE-B681-854DE5236146";

            var restoreCtxId = ApplicationContext.CurrentID;

            try
            {
                var sourceFileCase = (FileCase) commandActivity.WorkItem;

                if (sourceFileCase == null)
                    throw new ActivityCommandException("").WithUserMessage(LocalizedUserMessages.NoFile
                        .ToLocalizedName()
                        .FormatWith(commandActivity));

                if (!sourceFileCase.CanEdit(true))
                    throw new ActivityCommandException("").WithUserMessage(LocalizedUserMessages.NotEditable
                        .ToLocalizedName()
                        .FormatWith(sourceFileCase));

                //Parse tenant information from incoming
                string targetTenantId = String.Empty;

                var sourceIncoming = sourceFileCase.BaseFile.BaseIncomings
                        .Where(i => i.IncomingType.ToHasReferenceID().ReferenceID.ToUpper() ==
                                    incomingTypeVeReferenceId.ToUpper())
                        .OrderByDescending(t => t.CreatedAt)
                        .FirstOrDefault();
				if (sourceIncoming == null)
					throw new ActivityCommandException("").WithUserMessage(LocalizedUserMessages.IncomingNotFound
						.ToLocalizedName()
						.FormatWith(sourceFileCase.BaseFile));

				var targetFileCaseUri = new Uri(sourceIncoming.GetProperty("#LHIND_MitberichtExtern_SourceFileCaseUrl") as string);
				
				targetTenantId = HttpUtility.ParseQueryString(targetFileCaseUri.Query).Get("TenantID");

                var sourceFileCaseUrl = UrlProvider.Current.GetOpenWorkListItemUrl(sourceFileCase);


                //Create eCH-0147 container
                var messageExport = Containers.Global.Resolve<IMessageExport>();
                var eChExport = messageExport.Export(sourceFileCase);

                sourceFileCase.AddFileCaseContent(eChExport);

                //switch tenant
                using (ClientTransaction.CreateRootTransaction().EnterDiscardingScope())
                using (TenantSection.SwitchToTenant(Tenant.FindByUnqiueIdentifier(targetTenantId)))
                {
                    //Create new incoming, set specialdata properties
                    var incoming = Incoming.NewObject();
                    ApplicationContext.CurrentID = incoming.ApplicationContextID;

                    incoming.Subject = sourceFileCase.DisplayName + " - eCH Response";
                    incoming.IncomingType =
                        new ReferenceHandle<IncomingTypeClassificationType>(incomingTypeGsReferenceId)
                            .GetObject();
                    incoming.ExternalNumber = sourceFileCase.FormattedNumber;
                    incoming.Remark = sourceFileCase.WorkInstruction;
                    using (new SpecialdataIgnoreReadOnlySection())
                    {
                        incoming.SetProperty("#LHIND_MitberichtExtern_SourceFileCaseUrl", sourceFileCaseUrl);
                    }

                    var targeteChDocument = Document.NewObject(incoming);
                    ((IDocument) targeteChDocument).Name = sourceFileCase.GetMultilingualValue(fc => fc.Title) + " (" +
                                                           sourceFileCase.FormattedNumber + ") - eCH Import";
                    targeteChDocument.PhysicallyPresent = false;
                    targeteChDocument.Type =
                        (DocumentClassificationType) ClassificationType.GetObject(WellKnownObjects
                            .DocumentClassification.EchImport.GetObjectID());

                    using (SecurityFreeSection.Activate())
                    using (TenantSection.DisableQueryRestrictions())
                    using (var handle = eChExport.ActiveContent.GetContent())
                    {
                        targeteChDocument.ActiveContent.SetContent(handle, "zip", "application/zip");

                        var targetFile = ImportHelper.TenantKnowsObject(targeteChDocument, true);
                        if (targetFile != null)
                        {
                            incoming.LeadingGroup = targetFile.LeadingGroup;
                            incoming.Insert(targetFile);
                        }
                        else
                        {
                            incoming.LeadingGroup = UserHelper.Current.GetActaNovaUserExtension().StandardGroup != null
                                ? UserHelper.Current.GetActaNovaUserExtension().StandardGroup.AsTenantGroup()
                                : UserHelper.Current.OwningGroup.AsTenantGroup();
                        }
                    }

                    ClientTransaction.Current.Commit();
                }
            }
            catch (Exception ex)
            {
                s_logger.Error(ex.Message, ex);
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
            if (commandActivity.EffectiveWorkItemType.IsAssignableFrom(typeof(FileCase)))
                return null;
            return "Invalid WorkItemType";
        }
    } //End of LhindMitberichtReturnToGS

    internal static class ImportHelper
    {
        public static File TenantKnowsObject(Document eChDocument, bool autoImport = false)
        {
            using (var handle = eChDocument.ActiveContent.GetContent())
            using (var echZipReader = new StreamEchZipReader(handle.CreateStream()))
            {
                var message = echZipReader.ExtractMessage();
                if ((message != null) & (message.content != null) && message.content.dossiers != null &&
                    message.content.dossiers.dossier != null && message.content.dossiers.dossier.Count == 1)
                {
                    var file = MessageImportHelper.TryGetExistingDomainObject<File>(message.content.dossiers.dossier
                        .First()
                        .applicationCustom);
                    if (file != null && autoImport)
                    {
                        var messageImport = MessageImport.NewObject(
                            Containers.Global.Resolve<IMessageImportMapper>(),
                            echZipReader, MessageImportModeType.Values.CreateAndUpdate());
                        messageImport.ExecuteImport();
                    }

                    return file;
                }
            }

            return null;
        }
    }
}