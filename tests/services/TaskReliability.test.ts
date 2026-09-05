import { requestUrl } from "obsidian";
import {
	GoogleCalendarService,
	taskProjectionMatches,
} from "../../src/services/GoogleCalendarService";
import { TaskCalendarSyncService } from "../../src/services/TaskCalendarSyncService";

const UID = "61d3d239-e29c-4295-aac2-a40be5ace641";
const OTHER = "a0a38710-9c32-485a-bd39-9a049cc7e9ee";

describe("Google Calendar projection recovery", () => {
	function projectionFixture() {
		const task: any = {
			path: "Tasks/new-name.md",
			title: "Example",
			status: "ready",
			scheduled: "2026-09-06",
			archived: false,
		};
		const file: any = { path: task.path };
		const plugin: any = {
			app: {
				vault: {
					getName: () => "Example Vault",
					getMarkdownFiles: () => [file],
					read: async () =>
						`---\ntasknotesUid: ${UID}\ntags: ${task.archived ? "[archived]" : "[]"}\n---\n`,
				},
			},
			settings: {
				googleCalendarExport: {
					enabled: true,
					reconcileFromTasks: true,
					targetCalendarId: "test-calendar",
				},
			},
			fieldMapper: { mapFromFrontmatter: (fm: any) => ({ ...task, archived: Array.isArray(fm.tags) && fm.tags.includes("archived") }) },
			cacheManager: { getAllTasks: async () => [task], getTaskInfo: async () => task },
		};
		const owner = {
			tasknotesProjection: "1",
			tasknotesVault: "Example Vault",
			tasknotesUid: UID,
			tasknotesRole: "series",
		};
		const events: any[] = [
			{ id: "event1", etag: "v1", extendedProperties: { private: owner } },
		];
		const google: any = {
			listTaskProjections: jest.fn(async () => events),
			getConnectionGeneration: () => 0,
		};
		const service: any = new TaskCalendarSyncService(plugin, google);
		service.isEnabled = () => true;
		service.assertConnectionGenerationCurrent = async () => {};
		service.projectionIdentity = async () => UID;
		service.ownedProjectionIsCurrent = () => false;
		service.isTaskCalendarEligible = (t: any) => !t.archived;
		service.getTaskEventId = (t: any) => t.googleCalendarEventId;
		service.saveTaskEventId = jest.fn(async () => {});
		service.saveTaskExceptionMetadata = jest.fn(async () => {});
		service.syncTaskToCalendar = jest.fn(async () => true);
		service.deleteOrQueueCalendarEvent = jest.fn(async () => true);
		return { task, file, plugin, google, service, events, owner };
	}

	it("recovers a moved task's event link without a local provider index", async () => {
		const { task, service } = projectionFixture();
		await service.reconcileOwnedTaskProjections();
		expect(service.saveTaskEventId).toHaveBeenCalledWith(
			task.path,
			"event1",
			"test-calendar",
			0
		);
		expect(service.syncTaskToCalendar).toHaveBeenCalledTimes(1);
		expect(service.deleteOrQueueCalendarEvent).not.toHaveBeenCalled();
	});

	it("repairs a survivor before duplicate cleanup and keeps projections with missing sources", async () => {
		const { service, events, owner } = projectionFixture();
		events.push(
			{ id: "duplicate", extendedProperties: { private: owner } },
			{ id: "orphan", extendedProperties: { private: { ...owner, tasknotesUid: OTHER } } },
			{ id: "unowned", extendedProperties: { private: {} } }
		);
		await service.reconcileOwnedTaskProjections();
		expect(service.deleteOrQueueCalendarEvent.mock.calls.map((call: any[]) => call[2])).toEqual(
			["duplicate"]
		);
		expect(service.syncTaskToCalendar.mock.invocationCallOrder[0]).toBeLessThan(
			service.deleteOrQueueCalendarEvent.mock.invocationCallOrder[0]
		);
	});

	it("keeps duplicate events if the canonical survivor cannot be repaired", async () => {
		const { service, events, owner } = projectionFixture();
		events.push({ id: "duplicate", extendedProperties: { private: owner } });
		service.syncTaskToCalendar.mockResolvedValue(false);
		await service.reconcileOwnedTaskProjections();
		expect(service.deleteOrQueueCalendarEvent).not.toHaveBeenCalled();
	});

	it("deletes the projection of an archived task while retaining its identity", async () => {
		const { task, service } = projectionFixture();
		task.archived = true;
		await service.reconcileOwnedTaskProjections();
		expect(service.deleteOrQueueCalendarEvent).toHaveBeenCalledWith(
			task.path,
			"test-calendar",
			"event1"
		);
		expect(service.syncTaskToCalendar).not.toHaveBeenCalled();
	});

	it.each(["duplicate identity", "unreadable source", "attendees"])(
		"refuses cleanup on %s",
		async (kind) => {
			const { plugin, service, file, events } = projectionFixture();
			if (kind === "duplicate identity")
				plugin.app.vault.getMarkdownFiles = () => [file, { path: "Tasks/copy.md" }];
			if (kind === "unreadable source")
				plugin.app.vault.read = async () => {
					throw Error("unavailable");
				};
			if (kind === "attendees") events[0].attendees = [{ email: "invitee@example.com" }];
			await expect(service.reconcileOwnedTaskProjections()).rejects.toThrow();
			expect(service.saveTaskEventId).not.toHaveBeenCalled();
			expect(service.deleteOrQueueCalendarEvent).not.toHaveBeenCalled();
			expect(service.syncTaskToCalendar).not.toHaveBeenCalled();
		}
	);

	it.each([
		"missing source",
		"missing UUID",
		"unclosed frontmatter",
		"invalid YAML",
		"empty YAML",
	])("does not infer deletion from %s", async (kind) => {
		const { plugin, service } = projectionFixture();
		if (kind === "missing source") plugin.app.vault.getMarkdownFiles = () => [];
		else
			plugin.app.vault.read = async () =>
				({
					"missing UUID": "---\nstatus: ready\n---\n",
					"unclosed frontmatter": `---\ntasknotesUid: ${UID}\n`,
					"invalid YAML": `---\ntasknotesUid: [\n---\n`,
					"empty YAML": "---\n\n---\n",
				})[kind];
		await service.reconcileOwnedTaskProjections().catch(() => {});
		expect(service.deleteOrQueueCalendarEvent).not.toHaveBeenCalled();
	});

	function deletionFixture() {
		const fixture = projectionFixture();
		const { service, plugin, google } = fixture;
		plugin.settings.googleCalendarExport.syncOnTaskDelete = true;
		service.isDeletionQueueReady = () => true;
		service.withGoogleRateLimit = (fn: any) => fn();
		service.queueCalendarDeletion = jest.fn(async () => {});
		service.removeFromDeletionQueue = jest.fn(async () => {});
		google.deleteEvent = jest.fn(async () => {});
		delete service.deleteOrQueueCalendarEvent;
		return fixture;
	}

	it.each(["immediate", "queued"])(
		"keeps a damaged source safe through %s deletion",
		async (route) => {
			const { service, plugin, task, google } = deletionFixture();
			plugin.app.vault.read = async () => "---\nstatus: ready\n---\n";
			const item = {
				taskPath: task.path,
				calendarId: "test-calendar",
				eventId: "event1",
				attempts: 0,
				createdAt: 1,
			};
			if (route === "immediate") {
				expect(
					await service.deleteOrQueueCalendarEvent(task.path, "test-calendar", "event1")
				).toBe(false);
				expect(service.queueCalendarDeletion.mock.calls[0][3].message).toMatch(
					/stable identity/
				);
			} else {
				service.getDeletionQueue = async () => [item];
				service.mutateDeletionQueue = async (mutate: any) => mutate([item]);
				const result = await service.processDeletionQueue();
				expect(result).toEqual({ deleted: 0, failed: 1, remaining: 1 });
			}
			expect(google.deleteEvent).not.toHaveBeenCalled();
		}
	);

	it("requires a fresh archive tag rather than a cached archived flag", async () => {
		const { service, plugin, task, google } = deletionFixture();
		task.archived = true;
		plugin.app.vault.read = async () => `---\ntasknotesUid: ${UID}\n---\n`;
		expect(await service.deleteOrQueueCalendarEvent(task.path, "test-calendar", "event1")).toBe(
			false
		);
		expect(google.deleteEvent).not.toHaveBeenCalled();
	});

	it("deletes only a verified retirement, carrying its provider revision", async () => {
		const { service, task, google } = deletionFixture();
		task.archived = true;
		expect(await service.deleteOrQueueCalendarEvent(task.path, "test-calendar", "event1")).toBe(
			true
		);
		expect(google.deleteEvent).toHaveBeenCalledWith("test-calendar", "event1", 0, "v1");
	});

	it.each([false, true])(
		"deletes a duplicate only when all event content matches (changed=%s)",
		async (changed) => {
			const { service, task, google, events } = deletionFixture();
			task.googleCalendarEventId = "survivor";
			events.push({ ...events[0], id: "survivor", etag: "v2" });
			if (changed) events[0].attachments = [{ fileUrl: "https://example.com/attachment" }];
			expect(
				await service.deleteOrQueueCalendarEvent(task.path, "test-calendar", "event1")
			).toBe(!changed);
			expect(google.deleteEvent).toHaveBeenCalledTimes(changed ? 0 : 1);
		}
	);

	it("retires a detached occurrence only with an explicit completion or skip marker", async () => {
		const { service, task, events, google, owner } = deletionFixture();
		events[0].extendedProperties.private = {
			...owner,
			tasknotesRole: "exception",
			tasknotesOccurrence: "2026-09-06",
		};
		task.complete_instances = ["2026-09-06"];
		expect(await service.deleteOrQueueCalendarEvent(task.path, "test-calendar", "event1")).toBe(
			true
		);
		expect(google.deleteEvent).toHaveBeenCalledTimes(1);
	});

	it("recovers a detached recurrence event separately from its parent series", async () => {
		const { task, service, events, owner } = projectionFixture();
		task.googleCalendarExceptionOriginalScheduled = "2026-09-06";
		events.push({
			id: "detached",
			extendedProperties: {
				private: {
					...owner,
					tasknotesRole: "exception",
					tasknotesOccurrence: "2026-09-06",
				},
			},
		});
		await service.reconcileOwnedTaskProjections();
		expect(service.saveTaskExceptionMetadata).toHaveBeenCalledWith(
			task.path,
			{ googleCalendarExceptionEventId: "detached" },
			"test-calendar",
			0
		);
		expect(service.deleteOrQueueCalendarEvent).not.toHaveBeenCalled();
	});
});

describe("owned provider projection safety", () => {
	const owner = { tasknotesProjection: "1", tasknotesVault: "Example Vault", tasknotesUid: UID };
	function fixture() {
		const service: any = Object.create(GoogleCalendarService.prototype);
		service.baseUrl = "https://www.googleapis.com/calendar/v3";
		service.plugin = {
			app: { vault: { getName: () => "Example Vault" } },
			settings: {
				googleCalendarExport: {
					reconcileFromTasks: true,
					targetCalendarId: "test-calendar",
				},
			},
		};
		service.oauthService = { getValidToken: async () => "fixture-token" };
		service.withRetry = (fn: any) => fn();
		service.refreshAllCalendars = jest.fn(async () => {});
		service.convertToICSEvent = (event: any) => event;
		(requestUrl as jest.Mock).mockReset();
		return service;
	}
	it("pages through owned records, allowing a valid empty final page", async () => {
		const service = fixture();
		(requestUrl as jest.Mock)
			.mockResolvedValueOnce({
				json: {
					items: [{ id: "one" }, { id: "cancelled", status: "cancelled" }],
					nextPageToken: "next",
				},
			})
			.mockResolvedValueOnce({ json: {} });
		expect(await service.listTaskProjections("test-calendar", "Example Vault")).toEqual([
			{ id: "one" },
		]);
		const url = new URL((requestUrl as jest.Mock).mock.calls[1][0].url);
		expect(url.searchParams.getAll("privateExtendedProperty")).toEqual([
			"tasknotesProjection=1",
			"tasknotesVault=Example Vault",
		]);
		expect(url.searchParams.get("pageToken")).toBe("next");
	});

	it("refuses deletion if the provider changed after retirement was verified", async () => {
		const service = fixture();
		(requestUrl as jest.Mock).mockResolvedValue({
			json: { id: "event1", etag: "changed", extendedProperties: { private: owner } },
		});
		await expect(service.deleteEvent("test-calendar", "event1", 0, "verified")).rejects.toThrow(
			/changed/
		);
		expect((requestUrl as jest.Mock).mock.calls.map((call: any[]) => call[0].method)).toEqual([
			"GET",
		]);
	});

	it.each(["invitation", "unowned", "no etag"])(
		"refuses to delete a changed %s record",
		async (kind) => {
			const service = fixture();
			const event: any = {
				id: "fixture",
				etag: '"version1"',
				extendedProperties: { private: owner },
			};
			if (kind === "invitation") event.attendees = [{ email: "fixture@example.com" }];
			if (kind === "unowned") delete event.extendedProperties;
			if (kind === "no etag") delete event.etag;
			(requestUrl as jest.Mock).mockResolvedValue({ json: event });
			await expect(service.deleteEvent("test-calendar", "fixture")).rejects.toThrow();
			expect((requestUrl as jest.Mock).mock.calls.map((call) => call[0].method)).toEqual([
				"GET",
			]);
		}
	);
	it("uses the freshly read etag for deletion, including saved retries", async () => {
		const service = fixture();
		(requestUrl as jest.Mock)
			.mockResolvedValueOnce({
				json: { id: "fixture", etag: '"version1"', extendedProperties: { private: owner } },
			})
			.mockResolvedValueOnce({ status: 204 });
		await service.deleteEvent("test-calendar", "fixture");
		expect((requestUrl as jest.Mock).mock.calls[1][0]).toMatchObject({
			method: "DELETE",
			headers: { "If-Match": '"version1"' },
		});
	});
	it("detects remote edits even when the ownership marker is unchanged", () => {
		const expected = {
			summary: "Source",
			start: { date: "2026-09-05" },
			end: { date: "2026-09-06" },
			reminders: { useDefault: false },
			transparency: "transparent",
			visibility: "private",
			extendedProperties: { private: owner },
		};
		expect(taskProjectionMatches({ ...expected }, expected)).toBe(true);
		expect(taskProjectionMatches({ ...expected, summary: "Remote drift" }, expected)).toBe(
			false
		);
		expect(
			taskProjectionMatches({ ...expected, recurrence: ["RRULE:FREQ=DAILY"] }, expected)
		).toBe(false);
	});
});
